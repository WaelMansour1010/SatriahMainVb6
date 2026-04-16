Attribute VB_Name = "UpdateDatabase"
Dim sql As String
Dim i As Long
Dim s As String
Dim New_View  As String

Private Sub Update30Follow()
Dim s As String


If DB_CreateTable("TblRegDateDelgateDailsGrantee", True, "ID", True) = True Then

    '„›« ÌÕ Ê—»ÿ
    DB_CreateField "TblRegDateDelgateDailsGrantee", "DelgID", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Transaction_ID", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Transaction_Type", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "WarntID", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "ProjectID", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "ItemID", adInteger, adColNullable, , , ""

    '√⁄„œ… «·Ã—Ìœ
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Serial", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "MainID", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "MaDate", adDBTimeStamp, adColNullable, 8, , "", False
    DB_CreateField "TblRegDateDelgateDailsGrantee", "MainName", adVarWChar, adColNullable, 255, , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Interval", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Remarks", adVarWChar, adColNullable, 500, , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "StatusVisit", adInteger, adColNullable, , , ""

    ' ›«’Ì· ≈÷«›Ì… „‰ «·“Ì«—… «·√’·Ì…
    DB_CreateField "TblRegDateDelgateDailsGrantee", "DateOfRegularMaint", adDBTimeStamp, adColNullable, 8, , "", False
    DB_CreateField "TblRegDateDelgateDailsGrantee", "GranteeStartDate", adDBTimeStamp, adColNullable, 8, , "", False
    DB_CreateField "TblRegDateDelgateDailsGrantee", "GranteeEndDate", adDBTimeStamp, adColNullable, 8, , "", False
    DB_CreateField "TblRegDateDelgateDailsGrantee", "MaintenanceIDS", adInteger, adColNullable, , , ""
    DB_CreateField "TblRegDateDelgateDailsGrantee", "Done", adInteger, adColNullable, , , ""

End If



DB_CreateField "TBLRegularMaint", "StatusVisit", adInteger, adColNullable, , , ""


DB_CreateField "TblRegDateDelgate", "InvType", adInteger, adColNullable, , , ""

      If DB_CreateTable("tblLCOpenB", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
            DB_CreateField "tblLCOpenB", "ID", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "TblLCID", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "serial", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "MarginNo", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "GuaranteeDate", adDBTimeStamp, adColNullable, 8, , "", False
            DB_CreateField "tblLCOpenB", "Amount", adDouble, adColNullable, , , "    ", False, True
          DB_CreateField "tblLCOpenB", "AmountP", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblLCOpenB", "TotalAmount", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblLCOpenB", "ExpAmount", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblLCOpenB", "InsuranceAmount", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblLCOpenB", "PercentA", adDouble, adColNullable, , , "    ", False, True
            
            DB_CreateField "tblLCOpenB", "MarginAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
            DB_CreateField "tblLCOpenB", "BankAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
            
            DB_CreateField "tblLCOpenB", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
            DB_CreateField "tblLCOpenB", "NoteID2", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "NoteSerial2", adInteger, adColNullable, , , ""
            
            DB_CreateField "tblLCOpenB", "Type", adInteger, adColNullable, , , ""
            
            DB_CreateField "tblLCOpenB", "NoteID", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "NoteSerial", adInteger, adColNullable, , , ""
            
            DB_CreateField "tblLCOpenB", "NoteID2", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "NoteSerial2", adInteger, adColNullable, , , ""
            DB_CreateField "tblLCOpenB", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
            DB_CreateField "tblLCOpenB", "IsFullPayed", adBoolean, adColNullable, , , "", False, True
        
    
    End If
    DB_CreateField "TblAqarCommissions", "ValueAmount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblNotesSales", "ValueAmount", adDouble, adColNullable, , , "    ", False, True
    
    
    DB_CreateField "TblStore", "IsNotCreateEntry", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "TreatUncountedItemsAsZeroQty", adBoolean, adColNullable, , , "", False, True
    
    
    Dim MySQL As String

MySQL = ""
MySQL = MySQL & "SELECT " & vbCrLf
MySQL = MySQL & "    N.ChqueNum," & vbCrLf
MySQL = MySQL & "    N.ManualNo," & vbCrLf
MySQL = MySQL & "    DEV.Double_Entry_Vouchers_ID," & vbCrLf
MySQL = MySQL & "    DEV.Credit_Or_Debit," & vbCrLf
MySQL = MySQL & "    DEV.Value AS DEV_Value," & vbCrLf
MySQL = MySQL & "    DEV.RecordDateH," & vbCrLf
MySQL = MySQL & "    DEV.Account_Code," & vbCrLf
MySQL = MySQL & "    DEV.Double_Entry_Vouchers_Description AS DEV_DES," & vbCrLf
MySQL = MySQL & "    DEV.Double_Entry_Vouchers_Descriptione AS DevDESE," & vbCrLf
MySQL = MySQL & "    A.Account_Name," & vbCrLf
MySQL = MySQL & "    DEV.DEV_ID_Line_No," & vbCrLf
MySQL = MySQL & "    NT.NotesTypeName," & vbCrLf
MySQL = MySQL & "    DEV.UserID," & vbCrLf
MySQL = MySQL & "    U.UserName," & vbCrLf
MySQL = MySQL & "    DEV.RecordDate," & vbCrLf
MySQL = MySQL & "    DEV.Notes_ID," & vbCrLf
MySQL = MySQL & "    DEV.ReceiptID," & vbCrLf
MySQL = MySQL & "    DEV.OperaID," & vbCrLf
MySQL = MySQL & "    DEV.Transaction_ID," & vbCrLf
MySQL = MySQL & "    T.Transaction_Serial," & vbCrLf
MySQL = MySQL & "    T.Transaction_Date," & vbCrLf
MySQL = MySQL & "    TT.TransactionTypeName," & vbCrLf
MySQL = MySQL & "    DEV.PostedDate," & vbCrLf
MySQL = MySQL & "    DEV.PostedUserID," & vbCrLf
MySQL = MySQL & "    DEV.Account_Interval_ID," & vbCrLf
MySQL = MySQL & "    N.NoteDate," & vbCrLf
MySQL = MySQL & "    N.NoteType," & vbCrLf
MySQL = MySQL & "    N.NoteSerial," & vbCrLf
MySQL = MySQL & "    N.Note_Value," & vbCrLf
MySQL = MySQL & "    A.Account_Serial," & vbCrLf
MySQL = MySQL & "    A.Account_NameEng," & vbCrLf
MySQL = MySQL & "    A.Parent_Account_Code," & vbCrLf
MySQL = MySQL & "    A.opening_balance," & vbCrLf
MySQL = MySQL & "    A.opening_balance_type," & vbCrLf
MySQL = MySQL & "    A.Branch," & vbCrLf
MySQL = MySQL & "    A.Sum_account," & vbCrLf
MySQL = MySQL & "    A.cost_center," & vbCrLf
MySQL = MySQL & "    A.currenct_code," & vbCrLf
MySQL = MySQL & "    N.Remark," & vbCrLf
MySQL = MySQL & "    N.note_value_by_characters," & vbCrLf
MySQL = MySQL & "    N.foxy_no," & vbCrLf
MySQL = MySQL & "    DEV.DEV_ID_Line_No1," & vbCrLf
MySQL = MySQL & "    NT.NotesTypeNamee," & vbCrLf
MySQL = MySQL & "    TT.TransactionEnglishName," & vbCrLf
MySQL = MySQL & "    N.NoteSerial1," & vbCrLf
MySQL = MySQL & "    DEV.branch_id," & vbCrLf
MySQL = MySQL & "    B.ActivityTypeId," & vbCrLf
MySQL = MySQL & "    DEV.notes_all," & vbCrLf
MySQL = MySQL & "    B.branch_name," & vbCrLf
MySQL = MySQL & "    B.branch_namee," & vbCrLf
MySQL = MySQL & "    DEV.Posted," & vbCrLf
MySQL = MySQL & "    DEV.valuee AS DEV_ValueE," & vbCrLf
MySQL = MySQL & "    DEV.currency," & vbCrLf
MySQL = MySQL & "    DEV.rate," & vbCrLf
MySQL = MySQL & "    B.RegionID," & vbCrLf
MySQL = MySQL & "    S.name," & vbCrLf
MySQL = MySQL & "    S.namee," & vbCrLf
MySQL = MySQL & "    DEV.DescAccount," & vbCrLf
MySQL = MySQL & "    DEV.NextAccount_Code," & vbCrLf
MySQL = MySQL & "    DEV.project_id," & vbCrLf
MySQL = MySQL & "    DEV.opr_fullcode," & vbCrLf
MySQL = MySQL & "    DEV.projectid," & vbCrLf
MySQL = MySQL & "    DEV.IsHiddenInv," & vbCrLf
MySQL = MySQL & "    DEV.operid," & vbCrLf
MySQL = MySQL & "    DEV.pandid," & vbCrLf
MySQL = MySQL & "    DEV.Aqarid," & vbCrLf
MySQL = MySQL & "    AQ.aqarname," & vbCrLf
MySQL = MySQL & "    AQ.aqarNo," & vbCrLf
MySQL = MySQL & "    T.TaxFound," & vbCrLf
MySQL = MySQL & "    T.order_no,T.CBoBasedON " & vbCrLf

MySQL = MySQL & "FROM dbo.DOUBLE_ENTREY_VOUCHERS DEV" & vbCrLf
MySQL = MySQL & "INNER JOIN dbo.TblUsers U ON U.UserID = DEV.UserID" & vbCrLf
MySQL = MySQL & "INNER JOIN dbo.TblBranchesData B ON B.branch_id = DEV.branch_id" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.TblAqar AQ ON AQ.Aqarid = DEV.Aqarid" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.ACCOUNTS A ON A.Account_Code = DEV.Account_Code" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.TblSection S ON S.Id = B.RegionID" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.Notes N ON N.NoteID = DEV.Notes_ID" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.TblNotesTypes NT ON NT.NotesType = N.NoteType" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.Transactions T ON T.Transaction_ID = DEV.Transaction_ID" & vbCrLf
MySQL = MySQL & "LEFT JOIN dbo.TransactionTypes TT ON TT.Transaction_Type = T.Transaction_Type" & vbCrLf
MySQL = MySQL & "WHERE DEV.Posted IS NULL" & vbCrLf

Call db_createOrUpdateviewSQL("RptLedger_Sub", MySQL)



    DB_CreateField "markaas_taklefa", "akarid", adInteger, adColNullable, , , "  ", False, True
    
    If DB_CreateTable("TblExpensesDetVouch", True, "ID", True) = True Then
           DB_CreateField "TblExpensesDetVouch", "ExpID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpensesDetVouch", "ExpDetails", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpensesDetVouch", "CurrRow", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpensesDetVouch", "uintid", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpensesDetVouch", "Value", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "Rate", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "PriceTotal", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "vat", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "Vatyo", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "vaTotalPayedlue", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblExpensesDetVouch", "TotalPayed", adDouble, adColNullable, , , "    ", False, True
           
           DB_CreateField "TblExpensesDetVouch", "Unitss", adVarWChar, adColNullable, 4000, , "", False, True, , True
           DB_CreateField "TblExpensesDetVouch", "StrUnit", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "TblExpensesDetVouch", "type", adInteger, adColNullable, , , ""
         DB_CreateField "TblExpensesDetVouch", "iqarid", adInteger, adColNullable, , , ""
End If
  
    
     If DB_CreateTable("TblExpUnitNoVouch", True, "ID", True) = True Then
           DB_CreateField "TblExpUnitNoVouch", "ExpID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch", "ExpDetails", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch", "UnitID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch", "Valu", adDouble, adColNullable, , , "    ", False, True
           

End If
 
 
      If DB_CreateTable("TblExpUnitNoVouch1", True, "ID", True) = True Then
           DB_CreateField "TblExpUnitNoVouch1", "ExpID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch1", "ExpDetails", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch1", "UnitID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblExpUnitNoVouch1", "Valu", adDouble, adColNullable, , , "    ", False, True
           

End If
 

 DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "Unitss", adVarWChar, adColNullable, 4000, , "", False, True, , True
 DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "StrUnit", adVarWChar, adColNullable, 4000, , "", False, True, , True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "uintid", adInteger, adColNullable, , , ""
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "mType", adInteger, adColNullable, , , ""
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "iqarid", adInteger, adColNullable, , , ""
    
    
     DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "Unitss", adVarWChar, adColNullable, 4000, , "", False, True, , True
 DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "StrUnit", adVarWChar, adColNullable, 4000, , "", False, True, , True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "uintid", adInteger, adColNullable, , , ""
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "mType", adInteger, adColNullable, , , ""
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "iqarid", adInteger, adColNullable, , , ""
    
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "Aqarid", adInteger, adColNullable, , , ""
    
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "unittype", adInteger, adColNullable, , , ""
    
    
   
   DB_CreateField "TblOptions", "CostStartingGard", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "ShowPrinterDialoge2", adBoolean, adColNullable, , , "", False, True
    
    DB_CreateField "TblAging", "Emp_Id", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblAging", "Emp_Name", adVarWChar, adColNullable, 90, , "", False, True, , True
    
     If DB_CreateTable("tblGeneralCashingDetailsCus", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
            DB_CreateField "tblGeneralCashingDetailsCus", "ID", adInteger, adColNullable, , , ""
            DB_CreateField "tblGeneralCashingDetailsCus", "tblGeneralCashingId", adInteger, adColNullable, , , ""
            DB_CreateField "tblGeneralCashingDetailsCus", "CusId", adInteger, adColNullable, , , ""
            DB_CreateField "tblGeneralCashingDetailsCus", "Account_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
            DB_CreateField "tblGeneralCashingDetailsCus", "CusName", adVarWChar, adColNullable, 255, , "", False, True, , True
            DB_CreateField "tblGeneralCashingDetailsCus", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblGeneralCashingDetailsCus", "Vat", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblGeneralCashingDetailsCus", "Net", adDouble, adColNullable, , , "    ", False, True
            
    End If
    
    
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "IsExpens", adBoolean, adColNullable, , , "", False, True
    'DB_CreateField "notes_all", "ComResid", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "tblGeneralCashingDetailsCus", "Transaction_ID", adInteger, adColNullable, , , ""
    DB_CreateField "tblGeneralCashingDetailsCus", "NoteSerial1", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "TblOptions", "LimitDefaultCredit", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblOptions", "LimitDefaultCreditDays", adDouble, adColNullable, , , "    ", False, True
    
    
   DB_CreateField "Transactions", "CustomsReceiptDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "transactions", "CustomsValue", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "transactions", "AddValue", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "transactions", "TransactionStatus", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "transactions", "AddNotes", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "transactions", "Vendor", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "TblVATAvowal", "ActivityTypeId", adInteger, adColNullable, , , ""
    
    
    DB_CreateField "TblEmployee", "DateEndIndustrial", adDBTimeStamp, adColNullable, , , "   ", False, True
    'DB_CreateField "TblEmployee", "DateEndIndustrial", adDBTimeStamp, adColNullable, , , "   ", False, True
   DB_CreateField "TblEmployee", "DateEndIndustrialHijri", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "notes_all", "ComResid", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "notes_all", "NewNO", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPaymentType", "PaymentType", adInteger, adColNullable, , , ""
    
    DB_CreateField "Transactions", "IsPosted", adInteger, adColNullable, , , ""
    DB_CreateField "Transactions", "UserPosted", adInteger, adColNullable, , , ""
    
    
    DB_CreateField "notes_all", "Vendor", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "TblTravDueK", "Vendor", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblTravDueK", "ContractNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblTravDueK", "DuDate", adDBTimeStamp, adColNullable, , , "   ", False, True
    DB_CreateField "TblTravDueKDet", "ContItem", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblTravDueKDet", "PurchaseOrderNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblTravDueKDet", "LocationName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblTravDueKDet", "RentType", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "Notes", "UnitNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Notes", "ContItem", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Notes", "PurchaseOrderNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Notes", "LocationName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Notes", "RentType", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    
    DB_CreateField "notes", "FiterWaiver", adInteger, adColNullable, , , "        ", False, True
    DB_CreateField "TblFiterWaiver", "Discount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Notes", "FiterWaiverNoteSerial", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TblOtheExpensAqar", "Discount2", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblFiterWaiver", "FiterWaiver", adDouble, adColNullable, , , "    ", False, True
    
   DB_CreateField "TblFiterWaiver", "ForRenterB", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblOptions", "IsCahngeServiceInvoice", adBoolean, adColNullable, , , "        ", False, True
    
    DB_CreateField "TblOptions", "IsCreateOpenBalnceMan", adBoolean, adColNullable, , , "        ", False, True
    
    
    DB_CreateField "TblOrderUpload", "BillOfLadingNumber", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblEndDebtAgingInv", "BranchID", adInteger, adColNullable, , , "        ", False, True
    
    DB_CreateField "notes_all", "OrderMaintenanceId", adInteger, adColNullable, , , "        ", False, True
        
   DB_CreateField "TblCustemers", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "notes_all", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Transactions", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Notes", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Notes1", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    
    
    DB_CreateField "TblHandWages", "IsHiddenVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblHandWages", "SerPos", adInteger, adColNullable, , , "        ", False, True
    DB_CreateField "TblHandWages", "PaymentId", adInteger, adColNullable, , , "        ", False, True
    
DB_CreateField "tblContractInsAllocationsDetails", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "tblContractInsAllocationsDetails", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

DB_CreateField "Transactions", "ContainersCounts", adInteger, adColNullable, , , "        ", False, True
DB_CreateField "Transactions", "CarsCount", adInteger, adColNullable, , , "        ", False, True
   
    
    
   DB_CreateField "Accounts", "TblLCID", adInteger, adColNullable, , , ""
    
    DB_CreateField "Notes", "TblLCID", adInteger, adColNullable, , , ""
    DB_CreateField "Notes1", "TblLCID", adInteger, adColNullable, , , ""
    
    DB_CreateField "Notes1", "installIDCont", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "Notes1", "ManualNO", adInteger, adColNullable, , , ""

        DB_CreateField "TBLLCMargin2", "BankAccountCode2", adVarWChar, adColNullable, 255, , "", False, True, , True
     
    
 
    
    DB_CreateField "TBLLC", "NoteIDOpen", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLC", "NoteSerialOpen", adInteger, adColNullable, , , ""
    
    
    
    
   DB_CreateField "TBLLC", "AccountMarginParent", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TBLLC", "AccountLGParent", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TBLLC", "AccountAcceptanceParent", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TBLLC", "AccountExpensParent", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TBLLC", "AccountExpensCode", adVarWChar, adColNullable, 255, , "", False, True, , True
   
   DB_CreateField "TBLLC", "TypeLCLG", adInteger, adColNullable, , , ""
   
    DB_CreateField "TBLLC", "NoteID2", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLC", "NoteSerial2", adInteger, adColNullable, , , ""
    
    DB_CreateField "TBLLC", "LGExpPeriod", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLC", "LGExpiryDate", adDBTimeStamp, adColNullable, 8, , "", False
   DB_CreateField "FixedAssets", "Price", adDouble, adColNullable, , , "    ", False, True
    
         'DB_CreateField "TBLLC", "projectName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TBLLC", "MarginTotal2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItems", "IsPriceIsLenthWH", adBoolean, adColNullable, , , "  ", False, True
    DB_CreateField "TBLLC", "MarginTotal4", adDouble, adColNullable, , , "    ", False, True
    
        DB_CreateField "TBLLC", "MarginTotal3", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItems", "IsPriceIsLenthWH", adBoolean, adColNullable, , , "  ", False, True




DB_CreateField "notes_all", "PayAmount", adDouble, adColNullable, , , "    ", False, True


    
    DB_CreateField "Transactions", "poTransaction_ID", adInteger, adColNullable, , , ""
   DB_CreateField "Notes", "AcceptianPeriod", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLC", "AcceptianPeriod", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLC", "TotalBondHistory", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLC", "MarginTotal", adDouble, adColNullable, , , "    ", False, True
    
     DB_CreateField "TBLLC", "AccountExpProject", adVarWChar, adColNullable, 250, , "", False, True, , True
     
    DB_CreateField "LCTypes", "prifix", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TBLLC", "GuaranteeNo", adVarWChar, adColNullable, 255, , "", False, True, , True
     
     DB_CreateField "Projects", "NoteSerial", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "Projects", "NoteId", adInteger, adColNullable, , , ""
     DB_CreateField "TBLLC", "Account_Code2", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TBLLC", "Account_CodeExp", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TBLLC", "Account_CodeMargin", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TBLLC", "BondAmt", adDouble, adColNullable, , , "    ", False, True
     DB_CreateField "TBLLC", "GuaranteeDate", adDBTimeStamp, adColNullable, 8, , "", False
     DB_CreateField "TBLLC", "prifix", adVarWChar, adColNullable, 255, , "", False, True, , True
    
     
     DB_CreateField "TBLLC", "project_id", adInteger, adColNullable, , , ""
     
     DB_CreateField "branches", "a225", adVarWChar, adColNullable, 250, , "", False, True, , True
     
     DB_CreateField "branches", "a226", adVarWChar, adColNullable, 250, , "", False, True, , True
     DB_CreateField "branches", "a227", adVarWChar, adColNullable, 250, , "", False, True, , True
     DB_CreateField "BanksData", "PAcceptAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "BanksData", "PLCAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     
    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 6500 ,'⁄ﬁÊœ «·„ﬁ«Ê·Ì‰' ,'       Contract Invoice' ", "NotesType", 6500
    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 6501 ,'«·„‘«—Ì⁄' ,'       Project' ", "NotesType", 6501
    
        add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22001 ,'«·«⁄ „«œ« ' ,'       Project' ", "NotesType", 22001
        add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22002 ,'acceptant advice' ,'       Project' ", "NotesType", 22002
        add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22003 ,'acceptant advice' ,'       Project' ", "NotesType", 22003
     add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22005 ,'Close LC' ,'       Project' ", "NotesType", 22005
     
     DB_CreateField "BanksData", "PMarginAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     
     DB_CreateField "TBLLC", "AcceptAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TBLLC", "LCAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     DB_CreateField "TBLLC", "MarginAccount_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
     
        DB_CreateField "TblSalesPricesPlan", "IsNewPric", adInteger, adColNullable, , , ""
    
      DB_CreateField "TblSalesPricesPlan", "BoxId", adInteger, adColNullable, , , ""
       DB_CreateField "TblSalesPricesPlan", "FixedPOS", adBoolean, adColNullable, , , ""
       DB_CreateField "TblSalesPrices", "FixedPOS", adBoolean, adColNullable, , , ""
DB_CreateField "TblSalesPrices", "BoxId", adInteger, adColNullable, , , ""
DB_CreateField "Transaction_Details", "ProjectID", adInteger, adColNullable, , , "      ", False, True
DB_CreateField "Transactions", "ShipOrderNo", adVarWChar, adColNullable, 255, , "", False, True, , True
        
    DB_CreateField "branches", "T217", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T218", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T219", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T220", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T221", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T222", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T223", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "T224", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "projects", "Insurance", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "SubcontractorContract", "Insurance", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Transactions", "ShipOrderNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipEnquieryNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipAccountNo", adVarWChar, adColNullable, 255, , "", False, True, , True


DB_CreateField "tblContractInsAllocationsDetails", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True


DB_CreateField "tblContractInsAllocationsDetails", "Invoicetype", adInteger, adColNullable, , , "    ", False, True


DB_CreateField "TblCountriesGovernments", "Code", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "tblContractInsAllocationsDetails", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True



s = "CREATE OR ALTER VIEW RptLedger_sub_projects AS "
s = s & "SELECT dbo.projects.End_user_name, dbo.projects.sub_contractor_name, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.projects.total, "
s = s & "dbo.projects.sub_discount_total, dbo.projects.net, dbo.projects.items_total, dbo.RptLedger_Sub.Double_Entry_Vouchers_ID, dbo.RptLedger_Sub.Credit_Or_Debit, "
s = s & "dbo.RptLedger_Sub.DEV_Value, dbo.RptLedger_Sub.Account_Code, dbo.RptLedger_Sub.Account_Name, dbo.RptLedger_Sub.DEV_DES, dbo.RptLedger_Sub.DEV_ID_Line_No, "
s = s & "dbo.RptLedger_Sub.NotesTypeName, dbo.RptLedger_Sub.UserID, dbo.RptLedger_Sub.UserName, dbo.RptLedger_Sub.RecordDate, dbo.RptLedger_Sub.Notes_ID, "
s = s & "dbo.RptLedger_Sub.ReceiptID, dbo.RptLedger_Sub.Transaction_ID, dbo.RptLedger_Sub.OperaID, dbo.RptLedger_Sub.Transaction_serial, dbo.RptLedger_Sub.TransactionTypeName, "
s = s & "dbo.RptLedger_Sub.Transaction_Date, dbo.RptLedger_Sub.Posted, dbo.RptLedger_Sub.PostedDate, dbo.RptLedger_Sub.PostedUserID, dbo.RptLedger_Sub.Account_Interval_ID, "
s = s & "dbo.RptLedger_Sub.NoteSerial, dbo.RptLedger_Sub.NoteDate, dbo.RptLedger_Sub.NoteType, dbo.RptLedger_Sub.Note_Value, dbo.RptLedger_Sub.account_serial, "
s = s & "dbo.RptLedger_Sub.Account_NameEng, dbo.RptLedger_Sub.Parent_Account_Code, dbo.RptLedger_Sub.opening_balance, dbo.RptLedger_Sub.opening_balance_type, "
s = s & "dbo.RptLedger_Sub.Branch, dbo.RptLedger_Sub.Sum_account, dbo.RptLedger_Sub.cost_center, dbo.RptLedger_Sub.currenct_code, dbo.RptLedger_Sub.Remark, "
s = s & "dbo.RptLedger_Sub.note_value_by_characters, dbo.RptLedger_Sub.TransactionEnglishName, dbo.RptLedger_Sub.NotesTypeNameE, dbo.RptLedger_Sub.project_id, "
s = s & "dbo.RptLedger_Sub.DEV_ID_Line_No1, dbo.RptLedger_Sub.foxy_no, dbo.RptLedger_Sub.opr_fullcode, dbo.RptLedger_Sub.NoteSerial1, dbo.RptLedger_Sub.pandid, "
s = s & "dbo.RptLedger_Sub.operid, dbo.projects.Project_nameE, dbo.RptLedger_Sub.DevDESE, dbo.RptLedger_Sub.branch_namee, dbo.RptLedger_Sub.ManualNo, "
s = s & "dbo.projects.End_user_id, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee "
s = s & "FROM dbo.RptLedger_Sub INNER JOIN dbo.projects ON dbo.RptLedger_Sub.project_id = dbo.projects.id "
s = s & "LEFT OUTER JOIN dbo.TblCustemers ON dbo.projects.End_user_id = dbo.TblCustemers.CusID"
'Cn.Execute s

DB_CreateField "TblOptions", "ShowBalanceOfEmpInSalary", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "PaymentIntoAccouStat", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "AllowEditInvoiceNoticeDiscount", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "AllowEditInvoiceOfReturn", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "CloseMovingVchrinSales", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "IsMultiItemsInCompItem", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "CantChangeSalesPerson", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "BatchCreateManyworkOrder", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "CantChangeSalesPerson", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "ProvisionsByManagement", adBoolean, adColNullable, , , "", False, True




    DB_CreateField "TblOptions", "DiscountByQtyOnly", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "IsTransferByCode", adBoolean, adColNullable, , , "", False, True
    
    DB_CreateField "TblOptions", "ZacatHandW", adBoolean, adColNullable, , , "", False, True
    

DB_CreateField "tblActivitesType", "Company_Comment", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblActivitesType", "VATRegNo", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "StreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "AdditionalStreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "BuildingNumber", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "PlotIdentification", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "CityName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "PostalZone", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "CountrySubentity", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "CitySubdivisionName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "IdentificationCode", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "tblActivitesType", "POwithremainqty", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "tblActivitesType", "BankReturnID", adInteger, adColNullable, , , "    ", False, True

 DB_CreateField "tblActivitesType", "Commonname", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "SerialNumber", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "OrganizationName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 

  
  DB_CreateField "TblTripTypesTransport", "UnitId", adInteger, adColNullable, , , "    ", False, True
 DB_CreateField "tblActivitesType", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 DB_CreateField "tblActivitesType", "DefaultInvoicetype", adInteger, adColNullable, , , "    ", False, True
 
 DB_CreateField "TblStore", "BoxID", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "tblActivitesType", "SendingMode", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "tblActivitesType", "IsHiddenTransportInv", adBoolean, adColNullable, , , "", False, True

DB_CreateField "tblActivitesType", "industrey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "CSR", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "Privatekey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "PublickeycertPem", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 DB_CreateField "tblActivitesType", "SecretKey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblActivitesType", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "tblActivitesType", "ApplyEinvoice", adInteger, adColNullable, , , "    ", False, True



DB_CreateField "tblActivitesType", "ActivityName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
       DB_CreateField "tblActivitesType", "IsBluee", adBoolean, adColNullable, , , "", False, True
       
       
      
    DB_CreateField "tblActivitesType", "AllowScInterface2", adBoolean, adColNullable, , , "", False, True


'-------------------------------------------

        
   
    DB_CreateField "tblActivitesType", "Company_Arabic_Name", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "tblActivitesType", "Company_Name_Eng", adVarWChar, adColNullable, 255, , "", False, True, , True
 


    DB_CreateField "TblBranchesData", "Company_Arabic_Name", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblBranchesData", "Company_Name_Eng", adVarWChar, adColNullable, 255, , "", False, True, , True

DB_CreateField "TblBranchesData", "Company_Comment", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "TblBranchesData", "VATRegNo", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "StreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "AdditionalStreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "BuildingNumber", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "PlotIdentification", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "CityName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "PostalZone", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "CountrySubentity", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "CitySubdivisionName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "IdentificationCode", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblBranchesData", "POwithremainqty", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblBranchesData", "BankReturnID", adInteger, adColNullable, , , "    ", False, True

 DB_CreateField "TblBranchesData", "Commonname", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "SerialNumber", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "OrganizationName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 

  
 DB_CreateField "TblBranchesData", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 DB_CreateField "TblBranchesData", "DefaultInvoicetype", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "TblBranchesData", "SendingMode", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblBranchesData", "IsHiddenTransportInv", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblBranchesData", "industrey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "CSR", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "Privatekey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "PublickeycertPem", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 DB_CreateField "TblBranchesData", "SecretKey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "TblBranchesData", "ApplyEinvoice", adInteger, adColNullable, , , "    ", False, True



DB_CreateField "TblBranchesData", "Privatekey2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblBranchesData", "PublickeycertPem2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 DB_CreateField "TblBranchesData", "SecretKey2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True


DB_CreateField "TblOptions", "Privatekey2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "PublickeycertPem2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 DB_CreateField "TblOptions", "SecretKey2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True



DB_CreateField "TblBranchesData", "ActivityName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
       DB_CreateField "TblBranchesData", "IsBluee", adBoolean, adColNullable, , , "", False, True
       
       
      
    DB_CreateField "TblBranchesData", "AllowScInterface2", adBoolean, adColNullable, , , "", False, True







       DB_CreateField "TblOptions", "StreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "AdditionalStreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "BuildingNumber", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "PlotIdentification", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "CityName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "PostalZone", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "CountrySubentity", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "CitySubdivisionName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "IdentificationCode", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblOptions", "POwithremainqty", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblOptions", "BankReturnID", adInteger, adColNullable, , , "    ", False, True

 DB_CreateField "TblOptions", "Commonname", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "SerialNumber", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "OrganizationName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 

  
 DB_CreateField "TblOptions", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 DB_CreateField "TblOptions", "DefaultInvoicetype", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "TblOptions", "SendingMode", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "tblOPtions", "IsHiddenTransportInv", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblOptions", "industrey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "CSR", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "Privatekey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblOptions", "PublickeycertPem", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
 DB_CreateField "TblOptions", "SecretKey", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "Invoicetype", adInteger, adColNullable, , , "    ", False, True
 
DB_CreateField "TblOptions", "ApplyEinvoice", adInteger, adColNullable, , , "    ", False, True



DB_CreateField "TblOptions", "ActivityName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
       DB_CreateField "tblOPtions", "IsBluee", adBoolean, adColNullable, , , "", False, True
       
              DB_CreateField "tblOPtions", "ApplyEinvoiceWithActive", adBoolean, adColNullable, , , "", False, True
              DB_CreateField "tblOPtions", "ApplyEinvoiceWithBranch", adBoolean, adColNullable, , , "", False, True
              DB_CreateField "tblOPtions", "HiddenBalanceFromBox", adBoolean, adColNullable, , , "", False, True
              
              DB_CreateField "TblBranchesData", "ApplyEinvoiceWithBranch", adBoolean, adColNullable, , , "", False, True
                            DB_CreateField "tblOPtions", "EmpAccountByDep", adBoolean, adColNullable, , , "", False, True
       
       
      
    DB_CreateField "tblOPtions", "AllowScInterface2", adBoolean, adColNullable, , , "", False, True


  
    If DB_CreateTable("tmptblEInvoice", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        'DB_CreateField "tmptblEInvoice", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tmptblEInvoice", "InvoiceID", adInteger, adColNullable, , , ""
        DB_CreateField "tmptblEInvoice", "DefaultInvoicetype", adInteger, adColNullable, , , ""
      
       DB_CreateField "tmptblEInvoice", "zatcaStatus", adInteger, adColNullable, , , , False, True
        DB_CreateField "tmptblEInvoice", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "IssueDate", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tmptblEInvoice", "IssueTim", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tmptblEInvoice", "DocumentCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "TaxCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "StreetName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "BuildingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "CityName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "PostalZone", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "CitySubdivisionName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "RegistrationName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "CompanyID", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "ItemName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "Qty", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmptblEInvoice", "Price", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmptblEInvoice", "CoCRCode", adVarWChar, adColNullable, 400, , "", False, True, , True
       
        DB_CreateField "tmptblEInvoice", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmptblEInvoice", "VatValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmptblEInvoice", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "tmptblEInvoice", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
        DB_CreateField "tmptblEInvoice", "serial", adInteger, adColNullable, , , ""
        

 DB_CreateField "tmptblEInvoice", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tmptblEInvoice", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tmptblEInvoice", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
   
        
DB_CreateField "tmptblEInvoice", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "tmptblEInvoice", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tmptblEInvoice", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tmptblEInvoice", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tmptblEInvoice", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tmptblEInvoice", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "tmptblEInvoice", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

    
    
    
    End If
    
    

    
    
    DB_CreateField "tmptblEInvoice", "Export", adInteger, adColNullable, , , , False, True
        DB_CreateField "tblEInvoice", "Export", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "tmptblEInvoice", "Export", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "tblEInvoice2", "Export", adBoolean, adColNullable, , , "        ", False, True
    
    DB_CreateField "tblEInvoice2", "ManualInvoiceNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "tblEInvoice2", "IqarName", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "tblEInvoice", "ManualInvoiceNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "tblEInvoice", "IqarName", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "tmptblEInvoice", "ManualInvoiceNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "tmptblEInvoice", "IqarName", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    
   DB_CreateField "TblPrintBarCode", "ColorID", adVarWChar, adColNullable, 255, , "", False, True, , True

   DB_CreateField "TblPrintBarCode", "NoCount", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "Area", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "Length", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "Height", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "Width", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "IsExpirDate", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "L", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "W", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "TblPrintBarCode", "Remarks", adVarWChar, adColNullable, 255, , "", False, True, , True
   
    
       DB_CreateField "TblPrintBarCode", "UnitName", adVarWChar, adColNullable, 255, , "", False, True, , True
          DB_CreateField "TblPrintBarCode", "UnitNamee", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblPrintBarCode", "UnitID", adInteger, adColNullable, , , "  ", False, True
       DB_CreateField "TblPrintBarCode", "LineID", adInteger, adColNullable, , , "  ", False, True
 
    
    DB_CreateField "TblPrintBarCode", "ItemSize", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPrintBarCode", "ClassId", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPrintBarCode", "ColorName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPrintBarCode", "SizeName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPrintBarCode", "CusName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblPrintBarCode", "ItemSize", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    
   DB_CreateField "tmptblEInvoice", "TaxCategoryPercent", adDouble, adColNullable, , , , False, True
  DB_CreateField "tmptblEInvoice", "TaxCategoryID", adVarWChar, adColNullable, 10, , "", False, True, , True
   
   
DB_CreateField "tblEInvoice", "TaxCategoryID", adVarWChar, adColNullable, 10, , "", False, True, , True
    
     DB_CreateField "tblEInvoice2", "TaxCategoryPercent", adDouble, adColNullable, , , , False, True
  DB_CreateField "tblEInvoice2", "TaxCategoryID", adVarWChar, adColNullable, 10, , "", False, True, , True
    
    
    DB_CreateField "tmptblEInvoice", "ExcelFile", adVarWChar, adColNullable, 400, , "", False, True, , True
    DB_CreateField "tmptblEInvoice", "ExcelRow", adInteger, adColNullable, , , ""
    DB_CreateField "tmptblEInvoice", "Transaction_ID", adInteger, adColNullable, , , ""
   DB_CreateField "tmptblEInvoice", "InvoiceID", adInteger, adColNullable, , , ""
     
   DB_CreateField "tmptblEInvoice", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "tmptblEInvoice", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tmptblEInvoice", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True

DB_CreateField "tmptblEInvoice", "AdditionalStreetName", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "PlotIdentification", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "CountrySubentity", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tmptblEInvoice", "IdentificationCode", adVarWChar, adColNullable, 4000, , "", False, True, , True

 DB_CreateField "tmptblEInvoice", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True

   
   DB_CreateField "tmptblEInvoice", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "", False, True, , True
   
   
   DB_CreateField "tblEInvoice", "Identificationid", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "tmptblEInvoice", "Identificationid", adVarWChar, adColNullable, 255, , "", False, True, , True
   
   DB_CreateField "tblEInvoice", "schemeID", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "tmptblEInvoice", "schemeID", adVarWChar, adColNullable, 255, , "", False, True, , True
   


    
    DB_CreateField "tmptblEInvoice", "ComResid", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "tmptblEInvoice", "NewNO", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "tblEInvoice", "ComResid", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "tblEInvoice", "NewNO", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "tblEInvoice2", "ComResid", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "tblEInvoice2", "NewNO", adVarWChar, adColNullable, 255, , "", False, True, , True
    

DB_CreateField "tblEInvoice2", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "tblEInvoice", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "tmptblEInvoice", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

   DB_updateField "tblEInvoice", "Transaction_ID", "float not null  "
   DB_updateField "tblEInvoice2", "Transaction_ID", "float not null  "

DB_updateField "tmptblEInvoice", "Transaction_ID", "float not null  "



  DB_CreateField "tmptblEInvoice", "branch_id", adInteger, adColNullable, , , ""
  DB_CreateField "tblEInvoice", "branch_id", adInteger, adColNullable, , , ""
  DB_CreateField "tblEInvoice2", "branch_id", adInteger, adColNullable, , , ""
  
  DB_CreateField "tblEInvoice", "branch_name", adVarWChar, adColNullable, 255, , "", False, True, , True
  DB_CreateField "tblEInvoice2", "branch_name", adVarWChar, adColNullable, 255, , "", False, True, , True
  DB_CreateField "tmptblEInvoice", "branch_name", adVarWChar, adColNullable, 255, , "", False, True, , True

DB_CreateField "TblRegDateDelgate", "RowId", adGUID, adColNullable, , , "", False, True

DB_CreateField "tblEInvoice2", "GroupUniqueFileMaster", adGUID, adColNullable, , , "", False, True, , False
DB_CreateField "tblEInvoice", "GroupUniqueFileMaster", adGUID, adColNullable, , , "", False, True, , False
DB_CreateField "tmptblEInvoice", "GroupUniqueFileMaster", adGUID, adColNullable, , , "", False, True, , False


DB_CreateField "tblEInvoice", "Prefix", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice2", "GroupUniqueCode", adGUID, adColNullable, , , "", False, True, , False
DB_CreateField "tblEInvoice", "GroupUniqueCode", adGUID, adColNullable, , , "", False, True, , False
DB_CreateField "tmptblEInvoice", "GroupUniqueCode", adGUID, adColNullable, , , "", False, True, , False





 
DB_CreateField "tblEInvoice", "branchname", adVarWChar, adColNullable, 255, , "", False, True, , True
  DB_CreateField "tblEInvoice2", "branchname", adVarWChar, adColNullable, 255, , "", False, True, , True
  DB_CreateField "tmptblEInvoice", "branchname", adVarWChar, adColNullable, 255, , "", False, True, , True
  
DB_CreateField "tblEInvoice", "Export", adInteger, adColNullable, , , , False, True
DB_CreateField "tblEInvoice2", "Export", adInteger, adColNullable, , , , False, True

DB_updateField "tmptblEInvoice", "Transaction_ID", "float not null  "

'
'   DB_updateField "tmptblEInvoice", "InvoiceID", "nvarchar(40) not null  "
'   DB_updateField "tblEInvoice", "InvoiceID", "nvarchar(40) not null  "
'   DB_updateField "tblEInvoice2", "InvoiceID", "nvarchar(40) not null  "

  If DB_CreateTable("tblFromWeb", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "tblFromWeb", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tblFromWeb", "OrderID", adInteger, adColNullable, , , ""
        DB_CreateField "tblFromWeb", "OrderNo", adInteger, adColNullable, , , ""
        DB_CreateField "tblFromWeb", "TransType", adInteger, adColNullable, , , ""
        DB_CreateField "tblFromWeb", "Date", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblFromWeb", "StartDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblFromWeb", "EndDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblFromWeb", "FromTime", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblFromWeb", "ToTime", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblFromWeb", "ChkSallary", adBoolean, adColNullable, , , "  ", False, True
        DB_CreateField "tblFromWeb", "Code", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblFromWeb", "EmployeeCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblFromWeb", "EmployeeName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblFromWeb", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblFromWeb", "Notes", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblFromWeb", "Items", adVarWChar, adColNullable, 4000, , "", False, True, , True
        
        

        
        
    
    
    
    End If
    DB_CreateField "tblItems", "IsPriceIsMatrix", adBoolean, adColNullable, , , "  ", False, True
    
    ConvertInvoiceIdAll
    
    
         If DB_CreateTable("tblItemsMatrix", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
            DB_CreateField "tblItemsMatrix", "ID", adInteger, adColNullable, , , ""
            DB_CreateField "tblItemsMatrix", "ItemId", adInteger, adColNullable, , , ""
            DB_CreateField "tblItemsMatrix", "SerID", adInteger, adColNullable, , , ""
            DB_CreateField "tblItemsMatrix", "Value", adInteger, adColNullable, , , """"
                        
            
            Dim mText As Integer
            For i = 1 To 60
                If i = 1 Then
                    mText = 100
                End If
                
                DB_CreateField "tblItemsMatrix", "A" & CStr(mText), adDouble, adColNullable, , , "    ", False, True
                
                mText = mText + 50
                
            Next


        End If
        
        
        
        DB_CreateField "TblTripTypesTransport", "FromPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTripTypesTransport", "ToPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTripTypesTransport", "Price", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblTripTypesTransport", "FromCityID", adInteger, adColNullable, , , """"
        DB_CreateField "TblTripTypesTransport", "ToCityID", adInteger, adColNullable, , , """"
        
        
        DB_CreateField "TblClientTransContrDet", "FromPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblClientTransContrDet", "ToPrice", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblClientTransContrDet", "FromCityID", adInteger, adColNullable, , , """"
        DB_CreateField "TblClientTransContrDet", "ToCityID", adInteger, adColNullable, , , """"
        
        
        DB_CreateField "TblProjePayPrePayed", "NCashingType", adInteger, adColNullable, , , """"
        
        
 DB_CreateField "Transactions", "ShipCustomerName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipDistance", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipArea", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipSiteNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipProjectName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipStructuralElement", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipMixDescription", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipDriverName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipPipeLine", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipPump", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipTruckNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipIceTemp", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipTotalDeleveryd", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipThisLoad", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipDayOrder", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipTripNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipPlantNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipBatched", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "Transactions", "ShipRestunedPlant", adDBTimeStamp, adColNullable
    DB_CreateField "Transactions", "ShipEndDischarge", adDBTimeStamp, adColNullable
    DB_CreateField "Transactions", "ShipStartDisCharge", adDBTimeStamp, adColNullable
    DB_CreateField "Transactions", "ShipOnSite", adDBTimeStamp, adColNullable
    
    '****************************
    DB_CreateField "Tblposdata", "priceID", adInteger, adColNullable, , 0, "  ", False, True, , True
    '*****************
    DB_CreateField "TblItems", "ItemLimitType", adInteger, adColNullable, , 0, "  ", False, True, , True
    DB_CreateField "TblItems", "PeriodT1", adDouble, adColNullable, , , "    ", False, True
    '*****************
    '*****************
    DB_CreateField "notes_all", "ExcelFile", adVarWChar, adColNullable, 4000, "", "  ", False, True, , True
    DB_CreateField "notes_all", "ExcelRow", adInteger, adColNullable, , , 0, False, True
   
    DB_CreateField "TblItemShowDitailses", "TransType", adInteger, adColNullable, , 0, "    ", False, True
    DB_CreateField "TblItemShowDitailses", "Period", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "Plus", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "Min", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "PurQty", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "AvgQtyD", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "TotalQtyP", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShowDitailses", "ResultValue", adDouble, adColNullable, , , "    ", False, True
    '--------------------------------------------------
    DB_CreateField "TblItemShows", "TransType", adInteger, adColNullable, , 0, "    ", False, True
    DB_CreateField "TblItemShows", "Period", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "Plus", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "Min", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "PurQty", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "AvgQtyD", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "CusId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "TotalQtyP", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShows", "ResultValue", adDouble, adColNullable, , , "    ", False, True
    'DB_CreateField "TblItemShows", "CusId", adInteger, adColNullable, , , "    ", False, True
    '------------------------------------------------------
    
    DB_CreateField "TblItemGorupShowDitailses", "TransType", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemShInfo", "TransType", adInteger, adColNullable, , , "    ", False, True
    
    '-------------------
    DB_CreateField "TblItemShowBranch", "TransType", adInteger, adColFixed, , 0, "    ", False, True
'*****************************
    s = " SELECT dbo.Transaction_Details.Transaction_ID,"
    s = s & "        dbo.Transactions.Transaction_Date,"
    s = s & "        dbo.TblCustemers.CusName,"
    s = s & "        dbo.TblStore.StoreName,"
    s = s & "        dbo.Transaction_Details.Item_ID,"
    s = s & "        dbo.TblItems.ItemName,"
    s = s & "        dbo.TblItems.ItemCode,"
    s = s & "        dbo.Transaction_Details.ItemCase,"
    s = s & "        SUM(dbo.Transaction_Details.Quantity) AS Quantity,"
    s = s & "        dbo.Transaction_Details.Price,"
    s = s & "        dbo.Transaction_Details.ItemDiscountType,"
    s = s & "        dbo.Transaction_Details.ItemDiscount,"
    s = s & "        dbo.Transactions.Trans_Discount,"
    s = s & "        dbo.Transactions.Trans_DiscountType,"
    s = s & "        dbo.Transactions.TaxFound,"
    s = s & "        dbo.Transactions.TaxValue,"
    s = s & "        dbo.Transaction_Details.guaranteeTime,"
    s = s & "        dbo.Transactions.Transaction_Serial,"
    s = s & "        dbo.TblEmployee.Emp_Code,"
    s = s & "        dbo.TblEmployee.Emp_Name,"
    s = s & "        dbo.Transaction_Details.ItemSerial,"
    s = s & "        dbo.Transaction_Details.ShowQty,"
    s = s & "        dbo.Transaction_Details.showPrice,"
    s = s & "        dbo.TblUnites.UnitName,"
    s = s & "        dbo.Transaction_Details.Vat AS VatDet,"
    s = s & "        dbo.Transaction_Details.Vatyo,"
    s = s & "        dbo.Transaction_Details.MixNo,"
    s = s & "        dbo.Transaction_Details.QtyFaqtors,"
    s = s & "        dbo.Transaction_Details.FLgOrderSal,"
    s = s & "        dbo.Transactions.VAT,"
    s = s & "        dbo.Transactions.ResonVAT,"
    s = s & "        dbo.Transactions.Typ,"
    s = s & "        dbo.Transactions.VATNO,"
    s = s & "        dbo.Transactions.VATCustoms,"
    s = s & "        dbo.Transactions.VATCustoms1,"
    s = s & "        dbo.TblStore.StoreNamee,"
    s = s & "        dbo.Transactions.BLDate,"
    s = s & "        dbo.TblUnites.UnitNamee,"
    s = s & "        dbo.TblEmployee.Emp_Namee,"
    s = s & "        dbo.TblItems.ItemNamee,"
    s = s & "        dbo.TblCustemers.CusNamee,"
    s = s & "        dbo.Transactions.DueDate,"
    s = s & "        ISNULL(TblItems.ItemWithOutVAT, 0) ItemWithOutVAT"
    s = s & " From dbo.TblItems"
    s = s & "     INNER JOIN dbo.Transaction_Details"
    s = s & "         ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID"
    s = s & "     INNER JOIN dbo.TblStore"
    s = s & "         INNER JOIN dbo.TblCustemers"
    s = s & "             RIGHT OUTER JOIN dbo.Transactions"
    s = s & "                 ON dbo.TblCustemers.CusID = dbo.Transactions.CusID"
    s = s & "             ON dbo.TblStore.StoreID = dbo.Transactions.StoreID"
    s = s & "         ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "     INNER JOIN dbo.TblUnites"
    s = s & "         ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    s = s & "     LEFT OUTER JOIN dbo.TblEmployee"
    s = s & "         ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID"
    s = s & " GROUP BY dbo.Transaction_Details.Transaction_ID,"
    s = s & "          dbo.Transactions.Transaction_Date,"
    s = s & "          dbo.TblCustemers.CusName,"
    s = s & "          dbo.TblStore.StoreName,"
    s = s & "          dbo.Transaction_Details.Item_ID,"
    s = s & "          dbo.TblItems.ItemName,"
    s = s & "          dbo.TblItems.ItemCode,"
    s = s & "          dbo.Transaction_Details.ItemCase,"
    s = s & "          dbo.Transaction_Details.Price,"
    s = s & "          dbo.Transaction_Details.ItemDiscountType,"
    s = s & "          dbo.Transaction_Details.ItemDiscount,"
    s = s & "          dbo.Transactions.Trans_Discount,"
    s = s & "          dbo.Transactions.Trans_DiscountType,"
    s = s & "          dbo.Transactions.TaxFound,"
    s = s & "          dbo.Transactions.TaxValue,"
    s = s & "          dbo.Transaction_Details.guaranteeTime,"
    s = s & "          dbo.Transactions.Transaction_Serial,"
    s = s & "          dbo.TblEmployee.Emp_Code,"
    s = s & "          dbo.TblEmployee.Emp_Name,"
    s = s & "          dbo.Transaction_Details.ItemSerial,"
    s = s & "          dbo.Transaction_Details.ShowQty,"
    s = s & "          dbo.Transaction_Details.showPrice,"
    s = s & "          dbo.TblUnites.UnitName,"
    s = s & "          dbo.Transaction_Details.Vat,"
    s = s & "          dbo.Transaction_Details.Vatyo,"
    s = s & "          dbo.Transaction_Details.MixNo,"
    s = s & "          dbo.Transaction_Details.QtyFaqtors,"
    s = s & "          dbo.Transaction_Details.FLgOrderSal,"
    s = s & "          dbo.Transactions.VAT,"
    s = s & "          dbo.Transactions.ResonVAT,"
    s = s & "          dbo.Transactions.Typ,"
    s = s & "          dbo.Transactions.VATNO,"
    s = s & "          dbo.Transactions.VATCustoms,"
    s = s & "          dbo.Transactions.VATCustoms1,"
    s = s & "          dbo.TblStore.StoreNamee,"
    s = s & "          dbo.Transactions.BLDate,"
    s = s & "          dbo.TblUnites.UnitNamee,"
    s = s & "          dbo.TblEmployee.Emp_Namee,"
    s = s & "          dbo.TblItems.ItemNamee,"
    s = s & "          dbo.TblCustemers.CusNamee,"
    s = s & "          dbo.Transactions.DueDate,"
    s = s & "          TblItems.ItemWithOutVAT;"

    db_createOrUpdateviewSQL "QryBuyReportShort", s
    DB_CreateField "TBLCOUNTRIESDATA", "QRCODE", adVarWChar, adColNullable, 255, , "      ", False
   Dim rsDummy As New ADODB.Recordset
    s = "Select * from TBLCOUNTRIESDATA  Where IsNull(QRCODE,'') <> '' "
    rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly

    If rsDummy.EOF Then
        s = ""
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AD', 'Andorra', '√‰œÊ—«', 400);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AE', 'United Arab Emirates', '«·«„«—«  «·⁄—»Ì… «·„ Õœ…', 401);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AF', 'Afghanistan', '√›€«‰” «‰', 402);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AG', 'Antigua and Barbuda', '√‰ Ì€Ê« Ê»«—»Êœ«', 403);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AI', 'Anguilla', '√‰ÃÊÌ·«', 404);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AL', 'Albania', '√·»«‰Ì«', 405);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AM', 'Armenia', '√—„Ì‰Ì«', 406);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AO', 'Angola', '√‰ÃÊ·«', 407);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AQ', 'Antarctica', '«·ﬁ«—… «·ﬁÿ»Ì… «·Ã‰Ê»Ì…', 408);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AR', 'Argentina', '«·√—Ã‰ Ì‰', 409);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AS', 'American Samoa', '”«„Ê« «·√„—ÌﬂÌ…', 410);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AT', 'Austria', '«·‰„”«', 411);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AU', 'Australia', '√” —«·Ì«', 412);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AW', 'Aruba', '¬—Ê»«', 413);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AX', '?land Islands', 'Ã“— √Ê·«‰', 414);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('AZ', 'Azerbaijan', '√–—»ÌÃ«‰', 415);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BA', 'Bosnia and Herzegovina', '«·»Ê”‰… Ê«·Â—”ﬂ', 416);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BB', 'Barbados', '»—»«œÊ”', 417);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BD', 'Bangladesh', '»‰Ã·«œÌ‘', 418);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BE', 'Belgium', '»·ÃÌﬂ«', 419);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BF', 'Burkina Faso', '»Ê—ﬂÌ‰« ›«”Ê', 420);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BG', 'Bulgaria', '»·€«—Ì«', 421);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BH', 'Bahrain', '«·»Õ—Ì‰', 422);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BI', 'Burundi', '»Ê—Ê‰œÌ', 423);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BJ', 'Benin', '»‰Ì‰', 424);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BL', 'Saint Barthelemy', '”«‰ »«— Ì·„Ì', 425);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BM', 'Bermuda', '»—„Êœ«', 426);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BN', 'Brunei Darussalam', '»—Ê‰«Ì', 427);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BO', 'Bolivia (Plurinational State of)', '»Ê·Ì›Ì«', 428);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BQ', 'Bonaire, Sint Eustatius and Saba', '«·Ã“— «·ﬂ«—Ì»Ì… «·ÂÊ·‰œÌ…', 429);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BR', 'Brazil', '«·»—«“Ì·', 430);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BS', 'Bahamas', '«·»«Â«„«', 431);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BT', 'Bhutan', '»Ê «‰', 432);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BV', 'Bouvet Island', 'Ã“Ì—… »Ê›ÌÂ', 433);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BW', 'Botswana', '» ”Ê«‰«', 434);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BY', 'Belarus', '—Ê”Ì« «·»Ì÷«¡', 435);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('BZ', 'Belize', '»·Ì“', 436);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CA', 'Canada', 'ﬂ‰œ«', 437);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CC', 'Cocos (Keeling) Islands', 'Ã“— ﬂÊﬂÊ”', 438);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CD', 'Con (Democratic Republic of the)', 'Ã„ÂÊ—Ì… «·ﬂÊ‰€Ê «·œÌ„ﬁ—«ÿÌ…', 439);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CF', 'Central African Republic', 'Ã„ÂÊ—Ì… √›—ÌﬁÌ« «·Ê”ÿÏ', 440);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CG', 'Congo', 'Ã„ÂÊ—Ì… «·ﬂÊ‰€Ê', 441);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CH', 'Switzerland', '”ÊÌ”—«', 442);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CI', 'Cote díIvoire', '”«Õ· «·⁄«Ã', 443);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CK', 'Cook Islands', 'Ã“— ﬂÊﬂ', 444);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CL', 'Chile', '‘Ì·Ì', 445);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CM', 'Cameroon', '«·ﬂ«„Ì—Ê‰', 446);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CN', 'China', '«·’Ì‰', 447);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CO', 'Colombia', 'ﬂÊ·Ê„»Ì«', 448);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CR', 'Costa Rica', 'ﬂÊ” «—Ìﬂ«', 449);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CU', 'Cuba', 'ﬂÊ»«', 450);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CV', 'Cabo Verde', '«·—√” «·√Œ÷—', 451);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CW', 'CuraÁao', 'ﬂÊ—«”«Ê', 452);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CX', 'Christmas Island', 'Ã“Ì—… ⁄Ìœ «·„Ì·«œ', 453);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CY', 'Cyprus', 'ﬁ»—’', 454);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('CZ', 'Czechia', 'Ã„ÂÊ—Ì… «· ‘Ìﬂ', 455);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DE', 'Germany', '√·„«‰Ì«', 456);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DJ', 'Djibouti', 'ÃÌ»Ê Ì', 457);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DK', 'Denmark', '«·œ«‰„—ﬂ', 458);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DM', 'Dominica', 'œÊ„Ì‰Ìﬂ«', 459);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DO', 'Dominican Republic', 'Ã„ÂÊ—Ì… «·œÊ„Ì‰Ìﬂ«‰', 460);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('DZ', 'Algeria', '«·Ã“«∆—', 461);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('EC', 'Ecuador', '«·«ﬂÊ«œÊ—', 462);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('EE', 'Estonia', '«” Ê‰Ì«', 463);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('EG', 'Egypt', '„’—', 464);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('EH', 'Western Sahara', '«·’Õ—«¡ «·€—»Ì…', 465);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ER', 'Eritrea', '«—Ì —Ì«', 466);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ES', 'Spain', '√”»«‰Ì«', 467);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ET', 'Ethiopia', '«ÀÌÊ»Ì«', 468);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FI', 'Finland', '›‰·‰œ«', 469);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FJ', 'Fiji', '›ÌÃÌ', 470);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FK', 'Falkland Islands (Malvinas)', 'Ã“— ›Êﬂ·«‰œ', 471);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FM', 'Micronesia (Federated States of)', '„Ìﬂ—Ê‰Ì“Ì«', 472);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FO', 'Faroe Islands', 'Ã“— ›«—Ê', 473);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('FR', 'France', '›—‰”«', 474);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GA', 'Gabon', '«·Ã«»Ê‰', 475);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GB', 'United Kingdom of Great Britain and Northern Ireland', '«·„„·ﬂ… «·„ Õœ…', 476);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GD', 'Grenada', 'Ã—Ì‰«œ«', 477);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GE', 'Georgia', 'ÃÊ—ÃÌ«', 478);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GF', 'French Guiana', '€ÊÌ«‰«', 479);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GG', 'Guernsey', 'ÃÌ—‰“Ì', 480);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GH', 'Ghana', '€«‰«', 481);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GI', 'Gibraltar', 'Ã»· ÿ«—ﬁ', 482);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GL', 'Greenland', 'Ã—Ì‰·«‰œ', 483);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GM', 'Gambia', '€«„»Ì«', 484);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GN', 'Guinea', '€Ì‰Ì«', 485);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GP', 'Guadeloupe', 'ÃÊ«œ·Ê»', 486);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GQ', 'Equatorial Guinea', '€Ì‰Ì« «·«” Ê«∆Ì…', 487);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GR', 'Greece', '«·ÌÊ‰«‰', 488);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GS', 'South Georgia and the South Sandwich Islands', 'ÃÊ—ÃÌ« «·Ã‰Ê»Ì… ÊÃ“— ”«‰œÊÌ ‘ «·Ã‰Ê»Ì…', 489);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GT', 'Guatemala', 'ÃÊ« Ì„«·«', 490);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GU', 'Guam', 'ÃÊ«„', 491);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GW', 'Guinea-Bissau', '€Ì‰Ì« »Ì”«Ê', 492);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('GY', 'Guyana', '€Ì«‰«', 493);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HK', 'Hong Kong', 'ÂÊ‰€ ﬂÊ‰€', 494);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HM', 'Heard Island and McDonald Islands', 'Ã“Ì—… ÂÌ—œ ÊÃ“— „«ﬂœÊ‰«·œ', 495);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HN', 'Honduras', 'Â‰œÊ—«”', 496);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HR', 'Croatia', 'ﬂ—Ê« Ì«', 497);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HT', 'Haiti', 'Â«Ì Ì', 498);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('HU', 'Hungary', '«·„Ã—', 499);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ID', 'Indonesia', '«‰œÊ‰Ì”Ì«', 500);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IE', 'Ireland', '√Ì—·‰œ«', 501);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IL', 'Israel', '«”—«∆Ì·', 502);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IM', 'Isle of Man', 'Ã“Ì—… „«‰', 503);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IN', 'India', '«·Â‰œ', 504);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IO', 'British Indian Ocean Territory', '≈ﬁ·Ì„ «·„ÕÌÿ «·Â‰œÌ «·»—Ìÿ«‰Ì', 505);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IQ', 'Iraq', '«·⁄—«ﬁ', 506);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IR', 'Iran (Islamic Republic of)', '«Ì—«‰', 507);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IS', 'Iceland', '√Ì”·‰œ«', 508);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('IT', 'Italy', '«Ìÿ«·Ì«', 509);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('JE', 'Jersey', 'ÃÌ—”Ì', 510);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('JM', 'Jamaica', 'Ã«„«Ìﬂ«', 511);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('JO', 'Jordan', '«·√—œ‰', 512);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('JP', 'Japan', '«·Ì«»«‰', 513);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KE', 'Kenya', 'ﬂÌ‰Ì«', 514);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KG', 'Kyrgyzstan', 'ﬁ—€Ì“” «‰', 515);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KH', 'Cambodia', 'ﬂ„»ÊœÌ«', 516);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KI', 'Kiribati', 'ﬂÌ—Ì»« Ì', 517);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KM', 'Comoros', 'Ã“— «·ﬁ„—', 518);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KN', 'Saint Kitts and Nevis', '”«‰  ﬂÌ ” Ê‰Ì›Ì”', 519);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KP', 'Korea (Democratic Peopleís Republic of)', 'ﬂÊ—Ì« «·‘„«·Ì…', 520);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KR', 'Korea (Republic of)', 'ﬂÊ—Ì« «·Ã‰Ê»Ì…', 521);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KW', 'Kuwait', '«·ﬂÊÌ ', 522);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KY', 'Cayman Islands', 'Ã“— ﬂ«Ì„«‰', 523);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('KZ', 'Kazakhstan', 'ﬂ«“«Œ” «‰', 524);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LA', 'Lao Peopleís Democratic Republic', '·«Ê”', 525);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LB', 'Lebanon', '·»‰«‰', 526);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LC', 'Saint Lucia', '”«‰  ·Ê”Ì«', 527);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LI', 'Liechtenstein', '·ÌŒ ‰‘ «Ì‰', 528);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LK', 'Sri Lanka', '”—Ì·«‰ﬂ«', 529);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LR', 'Liberia', '·Ì»Ì—Ì«', 530);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LS', 'Lesotho', '·Ì”Ê Ê', 531);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LT', 'Lithuania', '·Ì Ê«‰Ì«', 532);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LU', 'Luxembourg', '·Êﬂ”„»Ê—Ã', 533);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LV', 'Latvia', '·« ›Ì«', 534);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('LY', 'Libya', '·Ì»Ì«', 535);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MA', 'Morocco', '«·„€—»', 536);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MC', 'Monaco', '„Ê‰«ﬂÊ', 537);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MD', 'Moldova (Republic of)', '„Ê·œ«›Ì«', 538);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ME', 'Montenegro', '«·Ã»· «·√”Êœ', 539);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MF', 'Saint Martin (French Part)', ' Ã„⁄ ”«‰ „«— Ì‰', 540);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MG', 'Madagascar', '„œ€‘ﬁ—', 541);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MH', 'Marshall Islands', 'Ã“— „«—‘«·', 542);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MK', 'North Macedonia', '„ﬁœÊ‰Ì«', 543);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ML', 'Mali', '„«·Ì', 544);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MM', 'Myanmar', '„Ì«‰„«—', 545);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MN', 'Mongolia', '„‰€Ê·Ì«', 546);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MO', 'Macao', '„«ﬂ«Ê', 547);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MP', 'Northern Mariana Islands', 'Ã“— „«—Ì«‰« «·‘„«·Ì…', 548);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MQ', 'Martinique', '„«— Ì‰Ìﬂ', 549);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MR', 'Mauritania', '„Ê—Ì «‰Ì«', 550);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MS', 'Montserrat', '„Ê‰ ”—« ', 551);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MT', 'Malta', '„«·ÿ«', 552);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MU', 'Mauritius', '„Ê—Ì‘ÌÊ”', 553);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MV', 'Maldives', 'Ã“— «·„«·œÌ›', 554);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MW', 'Malawi', '„·«ÊÌ', 555);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MX', 'Mexico', '«·„ﬂ”Ìﬂ', 556);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MY', 'Malaysia', '„«·Ì“Ì«', 557);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('MZ', 'Mozambique', '„Ê“„»Ìﬁ', 558);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NA', 'Namibia', '‰«„Ì»Ì«', 559);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NC', 'New Caledonia', 'ﬂ«·ÌœÊ‰Ì« «·ÃœÌœ…', 560);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NE', 'Niger', '«·‰ÌÃ—', 561);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NF', 'Norfolk Island', 'Ã“Ì—… ‰Ê—›Ê·ﬂ', 562);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NG', 'Nigeria', '‰ÌÃÌ—Ì«', 563);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NI', 'Nicaragua', '‰Ìﬂ«—«ÃÊ«', 564);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NL', 'Netherlands', 'ÂÊ·‰œ«', 565);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NO', 'Norway', '«·‰—ÊÌÃ', 566);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NP', 'Nepal', '‰Ì»«·', 567);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NR', 'Nauru', '‰Ê—Ê', 568);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NU', 'Niue', '‰ÌÊÌ', 569);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('NZ', 'New Zealand', '‰ÌÊ“Ì·«‰œ«', 570);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('OM', 'Oman', '⁄„«‰', 571);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PA', 'Panama', '»‰„«', 572);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PE', 'Peru', '»Ì—Ê', 573);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PF', 'French Polynesia', '»Ê·Ì‰“Ì« «·›—‰”Ì…', 574);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PG', 'Papua New Guinea', '»«»Ê« €Ì‰Ì« «·ÃœÌœ…', 575);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PH', 'Philippines', '«·›Ì·»Ì‰', 576);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PK', 'Pakistan', '»«ﬂ” «‰', 577);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PL', 'Poland', '»Ê·‰œ«', 578);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PM', 'Saint Pierre and Miquelon', '”«‰ »ÌÌ— Ê„Ìﬂ·Ê‰', 579);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PN', 'Pitcairn', '» ﬂ«Ì—‰', 580);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PR', 'Puerto Rico', '»Ê— Ê—ÌﬂÊ', 581);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PS', 'Palestinian, State of', '›·”ÿÌ‰', 582);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PT', 'Portugal', '«·»— €«·', 583);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PW', 'Palau', '»«·«Ê', 584);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('PY', 'Paraguay', '»«—«ÃÊ«Ì', 585);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('QA', 'Qatar', 'ﬁÿ—', 586);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('RE', 'RÈunion', '—ÊÌ‰ÌÊ‰', 587);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('RO', 'Romania', '—Ê„«‰Ì«', 588);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('RS', 'Serbia', '’—»Ì«', 589);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('RU', 'Russian Federation', '—Ê”Ì«', 590);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('RW', 'Rwanda', '—Ê«‰œ«', 591);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SA', 'Saudi Arabia', '«·”⁄ÊœÌ…', 592);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SB', 'Solomon Islands', 'Ã“— ”·Ì„«‰', 593);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SC', 'Seychelles', '”Ì‘·', 594);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SD', 'Sudan', '«·”Êœ«‰', 595);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SE', 'Sweden', '«·”ÊÌœ', 596);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SG', 'Singapore', '”‰€«›Ê—…', 597);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SH', 'Saint Helena, Ascension and Tristan da Cunha', '”«‰  ÂÌ·Ì‰« Ê√”Ì‰‘Ì‰ Ê —Ì” «‰ œ« ﬂÊ‰«', 598);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SI', 'Slovenia', '”·Ê›Ì‰Ì«', 599);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SJ', 'Svalbard and Jan Mayen', '”›«·»«—œ ÊÌ«‰ „«Ì‰', 600);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SK', 'Slovakia', '”·Ê›«ﬂÌ«', 601);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SL', 'Sierra Leone', '”Ì—«·ÌÊ‰', 602);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SM', 'San Marino', '”«‰ „«—Ì‰Ê', 603);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SN', 'Senegal', '«·”‰€«·', 604);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SO', 'Somalia', '«·’Ê„«·', 605);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SR', 'Suriname', '”Ê—Ì‰«„', 606);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SS', 'South Sudan', 'Ã‰Ê» «·”Êœ«‰', 607);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ST', 'Sao Tome and Principe', '”«Ê  Ê„Ì Ê»—Ì‰”Ì»', 608);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SV', 'El Salvador', '«·”·›«œÊ—', 609);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SX', 'Sint Maarten (Dutch Part)', '”Ì‰  „«— ‰', 610);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SY', 'Syrian Arab Republic', '”Ê—Ì«', 611);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('SZ', 'Eswatini', '”Ê«“Ì·«‰œ', 612);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TC', 'Turks and Caicos Islands', 'Ã“—  Ê—ﬂ” Êﬂ«ÌﬂÊ”', 613);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TD', 'Chad', ' ‘«œ', 614);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TF', 'French Southern Territories', '√—«÷ ›—‰”Ì… Ã‰Ê»Ì… Ê√‰ «— ÌﬂÌ…', 615);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TG', 'Togo', ' ÊÃÊ', 616);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TH', 'Thailand', ' «Ì·‰œ', 617);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TJ', 'Tajikistan', 'ÿ«Ãﬂ” «‰', 618);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TK', 'Tokelau', ' ÊﬂÌ·Ê', 619);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TL', 'Timor-Leste', ' Ì„Ê— «·‘—ﬁÌ…', 620);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TM', 'Turkmenistan', ' —ﬂ„«‰” «‰', 621);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TN', 'Tunisia', ' Ê‰”', 622);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TO', 'Tonga', ' Ê‰Ã«', 623);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TR', 'Turkey', ' —ﬂÌ«', 624);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TT', 'Trinidad and Tobago', ' —Ì‰Ìœ«œ Ê Ê»«€Ê', 625);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TV', 'Tuvalu', ' Ê›«·Ê', 626);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TW', 'Taiwan (Province of China)', ' «ÌÊ«‰', 627);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('TZ', 'Tanzania, United Republic of', ' «‰“«‰Ì«', 628);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('UA', 'Ukraine', '√Êﬂ—«‰Ì«', 629);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('UG', 'Uganda', '√Ê€‰œ«', 630);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('UM', 'United States Minor Outlying Islands', 'Ã“— «·Ê·«Ì«  «·„ Õœ… «·’€Ì—… «·‰«∆Ì…', 631);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('US', 'United States of America', '«·Ê·«Ì«  «·„ Õœ…', 632);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('UY', 'Uruguay', '√Ê—ÃÊ«Ì', 633);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('UZ', 'Uzbekistan', '√Ê“»ﬂ” «‰', 634);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VA', 'Holy See', '«·›« Ìﬂ«‰', 635);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VC', 'Saint Vincent and the Grenadines', '”«‰  ›Ì‰”‰  Ê«·€—Ì‰«œÌ‰', 636);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VE', 'Venezuela (Bolivarian Republic of)', '›‰“ÊÌ·«', 637);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VG', 'Virgin Islands (British)', 'Ã“— «·⁄–—«¡ «·»—Ìÿ«‰Ì…', 638);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VI', 'Virgin Islands (U.S.)', 'Ã“— «·⁄–—«¡ «·√„—ÌﬂÌ…', 639);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VN', 'Viet Nam', '›Ì ‰«„', 640);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('VU', 'Vanuatu', '›«‰Ê« Ê', 641);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('WF', 'Wallis and Futuna', 'Ê«·” Ê›Ê Ê‰«', 642);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('WS', 'Samoa', '”«„Ê«', 643);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('YE', 'Yemen', '«·Ì„‰', 644);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('YT', 'Mayotte', '„«ÌÊ ', 645);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ZA', 'South Africa', 'Ã‰Ê» √›—ÌﬁÌ«', 646);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ZM', 'Zambia', '“«„»Ì«', 647);"
        
        s = s & " INSERT INTO [TBLCOUNTRIESDATA]([QRCODE], [ECountryName], [CountryName], [CountryID])"
        s = s & " VALUES('ZW', 'Zimbabwe', '“Ì„»«»ÊÌ', 648);"
            
        Cn.Execute s
    End If
DB_CreateField "TBLTYPEVATS", "QRCODE", adVarWChar, adColNullable, 255, , "      ", False

            


            

        

    
    
     CurrentVersion = "V25-03-2026"  'lastlast 'lastlast 'lastlast 'lastlast 'lastlast
    

    'Cn.Execute "ALTER TABLE TblItems  DROP COLUMN ColorID"

    'updateNotesValueAndNobytext val(XPTxtID.Text), Format(XPTxtVal.Text, "###.00")
    ' adLongVarBinary Image
    'DB_CreateField "TblEmpData", "Photo2", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True

    'DB_updateField "notes_all", "foxy_no ", "float not null  "
    'DB_updateField "ACCOUNTS", "Account_Name", "nvarchar(4000)   "
    'update_record_to_table "TblNotesTypes", "NotesTypeNamee", " Account Opening Balance", "NotesType", 2000
    updateversion CurrentVersion, 30  ' funcid
    projectincludevchr
    ApprovalScreen
    UpdateDataBasePart29
    
    updateFuncSqaccountMovesl
    
    If SystemOptions.UserInterface = ArabicInterface Then

        MsgBox " „ «· ÕœÌÀ Õ Ï  «—ÌŒ  :" & CurrentVersion
    Else
        MsgBox "Data base Updated to Date : " & CurrentVersion
    End If

End Sub

Public Function UpdateDataBasePart28()
 
    Dim s As String

    '**************************
    DB_CreateField "projects_des", "qtySubContractor", adDouble, adColNullable, , , "", False, True
    DB_CreateField "projects_des", "costSubContractor", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "qtySubContractor", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "costSubContractor", adDouble, adColNullable, , , "", False, True
                  
    DB_CreateField "project_bill_details", "OLDTotalwithVat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "CurrenttotalWithvat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "Totalwitvat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "oldPerforValue", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "totalPerforValue", adDouble, adColNullable, , , "", False, True
    If DB_CreateTable("AqarRquestAndQutaion", True, "ID", True) = True Then
        DB_CreateField "AqarRquestAndQutaion", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
      
        DB_CreateField "AqarRquestAndQutaion", "Date", adDBTimeStamp, adColNullable
        DB_CreateField "AqarRquestAndQutaion", "BranchCode", adVarChar, adColNullable, 255, , , False
        DB_CreateField "AqarRquestAndQutaion", "Type", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "employee", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "CommissionType", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Section", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Value", adDouble, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "AkarUnit", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Country", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Schemes", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Aqar", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Governments", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "AkarUnit2", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Countrieshay", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Street", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "adress", adVarChar, adColNullable, 400, , , False
        DB_CreateField "AqarRquestAndQutaion", "Notes", adVarChar, adColNullable, 400, , , False
        DB_CreateField "AqarRquestAndQutaion", "GoogleLocation", adVarChar, adColNullable, 400, , , False
        DB_CreateField "AqarRquestAndQutaion", "FloorCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "DepartmentCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "StoreCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "RoomCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "OfficeCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Floor", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "BathRoomCount", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "BuildingYear", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Finished", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Name", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "ViewType", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "CustomerType", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "RecoredNo", adVarChar, adColNullable, 400, , , False
        DB_CreateField "AqarRquestAndQutaion", "CashType", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "UserId", adInteger, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "Tel1", adVarChar, adColNullable, 25, , , False
        DB_CreateField "AqarRquestAndQutaion", "Tel2", adVarChar, adColNullable, 25, , , False
        DB_CreateField "AqarRquestAndQutaion", "Tel3", adVarChar, adColNullable, 25, , , False
        DB_CreateField "AqarRquestAndQutaion", "HaseImage", adBoolean, adColNullable, , , , False
        DB_CreateField "AqarRquestAndQutaion", "HaseVedio", adBoolean, adColNullable, , , , False
    End If
    
        If DB_CreateTable("TblCaptinTrans2", True, "id ", True) = True Then
        
        DB_CreateField "TblCaptinTrans2", "MasterID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblCaptinTrans2", "Emp_ID", adInteger, adColNullable, , , " ", False, True
        DB_CreateField "TblCaptinTrans2", "CompanyName", adVarWChar, adColNullable, 4000, , " ", False, True, , True
        DB_CreateField "TblCaptinTrans2", "OperationName", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "TblCaptinTrans2", "EmpName", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "TblCaptinTrans2", "typename", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "TblCaptinTrans2", "Account_Name", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "TblCaptinTrans2", "DateEntry", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
        
        DB_CreateField "TblCaptinTrans2", "Amount", adDouble, adColNullable, , , "", False, True
        
    End If
    
    s = " SELECT YEAR,"
    s = s & "        Emp_ID,"
    s = s & "        t.emp_Name,"
    s = s & "        SUM(RetValue)     RetValue"
    s = s & " FROM   ("
    s = s & "            SELECT YEAR(dbo.Transactions.Transaction_Date) YEAR,"
    s = s & "                   dbo.Transactions.Emp_ID,"
    s = s & "                    te.Emp_Name,"
    s = s & "                   ("
    s = s & "                       SELECT SUM(Transaction_NetValue + ISNULL(vat, 0)) AS SumValue"
    s = s & "                       FROM   dbo.Transactions AS A"
    s = s & "                       Where (a.Transaction_Type = 9)"
    s = s & "                              AND (A.ReturnSerial = dbo.Transactions.NoteSerial1)"
    s = s & "                   ) AS RetValue"
    s = s & "            From dbo.transactions"
    s = s & "                   LEFT OUTER JOIN dbo.TblBranchesData"
    s = s & "                        ON  dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
    s = s & "                   LEFT OUTER JOIN dbo.TblEmployee AS te"
    s = s & "                        ON  dbo.Transactions.Emp_ID = te.Emp_ID"
    s = s & "            Where (dbo.transactions.PaymentType = 1)"
    s = s & "                   AND (dbo.Transactions.Transaction_Type = 21)"
    s = s & "        )                 T"
    s = s & " Where IsNull(retvalue, 0) <> 0"
    s = s & " Group By"
    s = s & "        YEAR,"
    s = s & "        Emp_ID,"
    s = s & "        t.emp_Name"

    db_createOrUpdateviewSQL "NetRemSalesMan", s



    s = " Create  FUNCTION [dbo].[GetBalanceQtyPO3] (@ItemID integer ,@order_no  nvarchar(255) ,@PurchaseNo  integer )"
    s = s & "  RETURNS Float"
    s = s & " AS"
    s = s & " Begin"
    s = s & " Return"
    
    s = s & " (SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ShowQty"
    s = s & "    FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
    s = s & "                  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "     Where dbo.transactions.Transaction_Type = 29"
    s = s & "                    AND (dbo.Transactions.NoteSerial1 = @order_no)  AND"
    s = s & "                   (dbo.Transaction_Details.Item_ID = @ItemID)  ) -"
    
    s = s & "  IsNull((SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ShowQty"
    s = s & "    FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
    s = s & "                   dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "     Where dbo.transactions.Transaction_Type = 22"
    s = s & "                    AND (dbo.Transactions.order_no = @order_no) AND (ISNULL(dbo.Transactions.CBoBasedON, 0) = 1) AND"
    s = s & "                   (dbo.Transaction_Details.Item_ID = @ItemID and Transactions.Transaction_ID <> @PurchaseNo)"
    s = s & "    )  ,0)"
    
    
    s = s & " End"

    db_createOrUpdateFuctionSQL "GetBalanceQtyPO3", s
 
 
    s = " SELECT YEAR,CusID,CusName,CusNamee,Sum(RetValue) RetValue  FROM ("
    s = s & " SELECT"
    s = s & "        Year(dbo.Transactions.Transaction_Date) Year,"
    s = s & "        dbo.Transactions.CusID,"
    s = s & "        dbo.TblCustemers.CusName,"
    s = s & "        dbo.TblCustemers.CusNamee,"

    s = s & "        ("
    s = s & "            SELECT SUM(Transaction_NetValue + ISNULL(vat, 0)) AS SumValue"
    s = s & "            FROM   dbo.Transactions AS A"
    s = s & "            Where (a.Transaction_Type = 9)"
    s = s & "                   AND (A.ReturnSerial = dbo.Transactions.NoteSerial1)"
    s = s & "        )  AS RetValue"

    '--  dbo.Transactions.Transaction_NetValue
    s = s & " From dbo.transactions"
    s = s & "        LEFT OUTER JOIN dbo.TblBranchesData"
    s = s & "             ON  dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
    s = s & "        LEFT OUTER JOIN dbo.TblCustemers"
    s = s & "             ON  dbo.Transactions.CusID = dbo.TblCustemers.CusID"
    s = s & " Where (dbo.transactions.PaymentType = 1)"
    s = s & "        AND (dbo.Transactions.Transaction_Type = 21)"

    s = s & " ) T"
    s = s & " Where IsNull(retvalue, 0) <> 0"
    s = s & " GROUP BY  Year,"
    s = s & "        CusID,"
    s = s & "        CusName,"
    s = s & "        CusNamee"

    db_createOrUpdateviewSQL "NetRemCus", s

    DB_CreateField "TblUsers", "OpenAtProduction", adBoolean, adColNullable, , , "", False, True

'    DB_CreateField "TblReCostCalcDet", "NoteSerial", adInteger, adColNullable, , , "      ", False, True
'    DB_CreateField "TblReCostCalcDet", "NoteID", adInteger, adColNullable, , , "      ", False, True
'    DB_CreateField "TblReCostCalcDet", "BranchId", adInteger, adColNullable, , , "      ", False, True

    
    DB_CreateField "opr_employee_details", "FromProjectID", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblVocationEntitlements", "PaymentRecommended", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblOptions", "LinkSupplerWithItem", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblUsers", "HideInfroCasher", adBoolean, adColNullable, , , "", False, True



    DB_CreateField "TblOptions", "ShowOnlyItemsOfSales", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblAqarDetai", "isTax", adBoolean, adColNullable, , , " ", False, True
    DB_CreateField "branches", "a207", adVarWChar, adColNullable, 50, , "", False, True, , True


    s = " SELECT DISTINCT"
    s = s & "    dbo.TblFiterWaiver.ID,"
    s = s & "        dbo.TblFiterWaiver.RecordDateH,"
    s = s & "        dbo.TblFiterWaiver.RecordDate,"
    s = s & "        dbo.TblFiterWaiver.BranchID,"
    s = s & "        dbo.TblFiterWaiver.BulidID,"
    s = s & "        dbo.TblAqar.aqarname,"
    s = s & "        dbo.TblFiterWaiver.RenterID,"
    s = s & "        dbo.TblCustemers.CusName,"
    s = s & "        dbo.TblCustemers.CusNamee,"
    s = s & "        dbo.TblFiterWaiver.ApartmentID,"
    s = s & "        dbo.TblAqarDetai.unitno,"
    s = s & "        dbo.TblFiterWaiver.EndDateH,"
    s = s & "        dbo.TblFiterWaiver.EndDate,"
    s = s & "        dbo.TblFiterWaiver.FilterDate,"
    s = s & "        dbo.TblFiterWaiver.FilterDateH,"
    s = s & "        t2.BillPrice,"
    s = s & "        dbo.TblFiterWaiver.AccountNo,"
    s = s & "        dbo.TblFiterWaiver.AmountDely,"
    s = s & "        dbo.TblFiterWaiver.DayNo,"
    s = s & "        dbo.TblFiterWaiver.UserID,"
    s = s & "        dbo.TblFiterWaiver.OFRenter,"
    s = s & "        dbo.TblFiterWaiver.ForRenter,"
    s = s & "        dbo.TblFiterWaiver.unittype,"
    s = s & "        dbo.TblAkarUnit.name         AS nameUnt,"
    s = s & "        dbo.TblAkarUnit.namee,"
    s = s & "        dbo.TblFiterWaiver.ContNo,"
    s = s & "        dbo.TblFiterWaiver.ContractNo,"
    s = s & "        dbo.TblFiterWaiver.NoteID,"
    s = s & "        dbo.TblFiterWaiver.NoteSerial,"
    s = s & "        dbo.TblFiterWaiver.ContractDays,"
    s = s & "        dbo.TblFiterWaiver.WaterPrice,"
    s = s & "        dbo.TblFiterWaiver.ActualDays,"
    s = s & "        dbo.TblFiterWaiver.DayPricen,"
    s = s & "        T2.WaterPriceotal,"
    s = s & "        T2.ServicePrice,"
    s = s & "        T2.DayPricentotal,"
    s = s & "        T2.Service,"
    s = s & "        T2.WaterPayed,"
    s = s & "        T2.RentValuePayed,"
    s = s & "        T2.OldRent TelandNetPayed,"
    s = s & "        T2.RemainWater,"
    s = s & "        T2.RemainRent,"
    s = s & "        T2.RemainService,"
    s = s & "        T2.Insurance,"
    s = s & "        T2.outflow,"
    s = s & "        T2.StartDate,"
    s = s & "        T2.StartDateh,"
    s = s & "        T2.TotalStill,"
    s = s & "        T2.RemainCommissions,"
    s = s & "        T2.NoDaye,"
    s = s & "        dbo.TblFiterWaiver.outCondition,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncrease,"
    s = s & "        dbo.TblFiterWaiver.DaysValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayValueInc,"
    s = s & "        dbo.TblFiterWaiver.DayCountInc,"
    s = s & "        dbo.TblFiterWaiver.DayValueIncomplete,"
    s = s & "        dbo.TblFiterWaiver.DayCountIncomplete,"
    s = s & "        dbo.TblFiterWaiver.Efflux,"
    s = s & "        dbo.TblFiterWaiver.ValDay,"
    s = s & "        dbo.TblFiterWaiver.Discount,"
    s = s & "        dbo.TblFiterWaiver.totalcollected,"
    s = s & "        dbo.TblFiterWaiver.totalpayed,"
    s = s & "        dbo.TblFiterWaiver.LegalIssue,"
    s = s & "        dbo.TblFiterWaiver.net"
    s = s & " From dbo.TblAkarUnit"
    s = s & "        RIGHT OUTER JOIN dbo.TblFiterWaiver"
    s = s & "             ON  dbo.TblAkarUnit.id = dbo.TblFiterWaiver.unittype"
    s = s & "        LEFT OUTER JOIN dbo.TblAqarDetai"
    s = s & "             ON  dbo.TblFiterWaiver.ApartmentID = dbo.TblAqarDetai.Id"
    s = s & "        LEFT OUTER JOIN dbo.TblCustemers"
    s = s & "             ON  dbo.TblFiterWaiver.RenterID = dbo.TblCustemers.CusID"
    s = s & "        LEFT OUTER JOIN dbo.TblAqar"
    s = s & "             ON  dbo.TblFiterWaiver.BulidID = dbo.TblAqar.Aqarid"
            
    s = s & "        LEFT OUTER JOIN dbo.TblFiterWaiverDet2 T2"
    s = s & "             ON  T2.MasterID = dbo.TblFiterWaiver.ID"
    
    db_createOrUpdateviewSQL "VwFiterWaiver", s

    DB_CreateField "TblFiterWaiver", "DaysValueIncrease", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DaysValueIncomplete", adDouble, adColNullable, , , " ", False, True

    DB_CreateField "TblFiterWaiver", "DayValueInc", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayCountInc", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayValueIncomplete", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayCountIncomplete", adDouble, adColNullable, , , " ", False, True

    DB_CreateField "TblItemsParts", "isPrinted", adBoolean, adColNullable, , , " ", False, True
              
    DB_CreateField "TblItems", "PrintedName", adVarWChar, adColNullable, 255, , "  ", False, True, , True

    If DB_CreateTable("TblFiterWaiverDet2", True, "ID", True) = True Then
        DB_CreateField "TblFiterWaiverDet2", "MasterID", adInteger, adColNullable, , , " ???    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "RecordDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFiterWaiverDet2", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "BulidID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RenterID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ApartmentID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "Insurance", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "EndDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "EndDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFiterWaiverDet2", "FilterDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "FilterDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFiterWaiverDet2", "BillPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "AccountNo", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "DayNo", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "AmountDely", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "DayLate", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "AmountDely", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "DaysValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "TotalDept", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "TotalRight", adDouble, adColNullable, , , "    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "unittype", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "OFRenter", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ForRenter", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RecordDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFiterWaiverDet2", "UserID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ContNo", adInteger, adColNullable, , , " ???    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "ContractNo", adVarWChar, adColNullable, 4000

        DB_CreateField "TblFiterWaiverDet2", "NoteID", adInteger, adColNullable, , , " ???    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True

        DB_CreateField "TblFiterWaiverDet2", "ContractDays", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ActualDays", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "WaterPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "DayPricen", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ServicePrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "WaterPriceotal", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "DayPricentotal", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "Service", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "WaterPayed", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RentValuePayed", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "TelandNetPayed", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RemainWater", adDouble, adColNullable, , , "    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "RemainWater", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RemainRent", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RemainService", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "outflow", adBoolean, adColNullable, , , "        ", False, True
        '18022015

        DB_CreateField "TblFiterWaiverDet2", "outCondition", adBoolean, adColNullable, , , "        ", False, True

        DB_CreateField "TblFiterWaiverDet2", "NoDaye", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "Efflux", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "ValDay", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "Discount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "totalcollected", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "totalpayed", adDouble, adColNullable, , , "    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "net", adDouble, adColNullable, , , "    ", False, True

        DB_CreateField "TblFiterWaiverDet2", "StartDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFiterWaiverDet2", "StartDateh", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblFiterWaiverDet2", "OldRent", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "RemainDays", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblFiterWaiverDet2", "DaysValue", adDouble, adColNullable, , , "    ", False, True

    End If

    DB_CreateField "TblFiterWaiverDet2", "TotalStill", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiverDet2", "RemainCommissions", adDouble, adColNullable, , , " ", False, True

End Function

Public Function UpdateDataBasePart24()

On Error Resume Next

DB_CreateField "TblPrintBarCode", "ProductionDate", adVarWChar, adColNullable, 50, , "      ", False, True, , True

If DB_CreateTable("TblNotesOwnerPayment202", True, "ID", True) = True Then
           DB_CreateField "TblNotesOwnerPayment202", "NoteID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "NoteID2", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "CusID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "Aqarid", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "UnitNo", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "branch_no", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "TypTrans", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "Remarks", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
           DB_CreateField "TblNotesOwnerPayment202", "ContNoteSerial1", adVarWChar, adColNullable, 255, , "      ", False, True, , True
           DB_CreateField "TblNotesOwnerPayment202", "NoteDate", adDBTimeStamp, adColNullable, , , "      ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "value", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "PayedValue", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "RemainingValue", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "NetValue", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblNotesOwnerPayment202", "TransPayedValue", adDouble, adColNullable, , , "    ", False, True
End If
 If DB_CreateTable("TblOwnerPayment202", True, "ID", True) = True Then
           DB_CreateField "TblOwnerPayment202", "NoteID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblOwnerPayment202", "NoteID2", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblOwnerPayment202", "TypTrans", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "TblOwnerPayment202", "value", adDouble, adColNullable, , , "    ", False, True
           DB_CreateField "TblOwnerPayment202", "PayedValue", adDouble, adColNullable, , , "    ", False, True

End If
DB_CreateField "TblNotesOwnerPayment202", "NoteID3", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblOwnerPayment202", "NoteID3", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblOwnerPayment202", "UnitNo", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "Notes", "TotalPayed", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblExpensesDet", "vaTotalPayedlue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblExpUnitNo", "TotalPayed", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "OfficeValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "RenterValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "ExpValue", adDouble, adColNullable, , , "    ", False, True

    
    s = " SELECT"
s = s & "        TblFiterWaiverDe.IDFItWaiv,TblAqar.aqarname,"
s = s & "        dbo.TblFiterWaiverDe.Remark,"
s = s & "        Sum(dbo.TblFiterWaiverDe.Price * dbo.TblFiterWaiverDe.[Count]) AS Total ,"
s = s & "        dbo.TblAqrCompenetDet.Name AS nameDet"
       
s = s & " From dbo.TblFiterWaiverDe"
s = s & "        LEFT OUTER JOIN dbo.TblAqrCompenetDet"
s = s & "             ON  dbo.TblFiterWaiverDe.IDItem = dbo.TblAqrCompenetDet.ID"
s = s & "        LEFT OUTER JOIN dbo.TblAqrCompenet"
s = s & "             ON  dbo.TblFiterWaiverDe.GroupID = dbo.TblAqrCompenet.ID"
            
s = s & "             LEFT OUTER JOIN tblFiterWaiver ON TblFiterWaiverDe.IDFItWaiv =tblFiterWaiver.ID"
s = s & "             LEFT OUTER JOIN TblAqar ON TblAqar.Aqarid =  tblFiterWaiver.BulidID"

s = s & " Where count <> 0"
s = s & " GROUP BY dbo.TblFiterWaiverDe.Remark,dbo.TblAqrCompenetDet.Name,TblFiterWaiverDe.IDFItWaiv,TblAqar.aqarname"
    

db_createOrUpdateviewSQL "View_WaiverExpens", s


       s = " SELECT   TotalExp = (SELECT SUM(COUNT * Price) TotalExpe FROM TblFiterWaiverDe),"
s = s & "       CountContract = (SELECT COUNT(*) FROM TblContract ), "
s = s & "                          Cashing = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 0 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               ),"
s = s & "               Commission        = SUM("
s = s & "                   Case Notes.CashingType"
s = s & "                        WHEN 12 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               ),"
s = s & "               Arbon             = SUM(CASE Notes.CashingType WHEN 9 THEN (Note_Value) ELSE 0 END),"
s = s & "               ValueTransfer     = SUM("
s = s & "                   Case Notes.NoteCashingType"
s = s & "                        WHEN 3 THEN (Note_Value)"
s = s & "                        ELSE 0"
s = s & "                   End"
s = s & "               )"
s = s & "        From Notes "
db_createOrUpdateviewSQL "View_TotalsIq", s
Dim MySQL As String
MySQL = " SELECT     dbo.Notes.NoteID,tu.UserName, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.NoteDateH,"
MySQL = MySQL & "                       dbo.Notes.ContractNo, dbo.Notes.ContNo, dbo.Notes.commission, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.FilterID, dbo.Notes.FIlterTotal, dbo.Notes.Instrunce,"
MySQL = MySQL & "                       dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.NoteOrBonID, dbo.Notes.comXold, dbo.Notes.ComYold, dbo.Notes.NoteOrBonValue,"
MySQL = MySQL & "                       dbo.Notes.NoteOrBonSereal, dbo.Notes.Telephone, dbo.Notes.CashingType, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "                       dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Notes.renterName, dbo.Notes.NoteCashingType, dbo.Notes.BankName, dbo.Notes.DueDate,"
MySQL = MySQL & "                       dbo.Notes.ChqueNum, dbo.Notes.Remark, dbo.Notes.Remark2, dbo.Notes.ToPriodDateH, dbo.Notes.FrmPriodDateH, dbo.Notes.ToPriodDate, dbo.Notes.FrmPriodDate,"
MySQL = MySQL & "                       dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.unitno,"
MySQL = MySQL & "                       dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.Aqarid, TblAqar_1.aqarname, TblAkarUnit_2.name, TblAkarUnit_2.namee, dbo.Notes.akarid,"
                      MySQL = MySQL & " TblAqar_1.aqarname AS aqarname2, dbo.Notes.unittype AS unittype2, TblAkarUnit_1.name AS name2, TblAkarUnit_1.namee AS namee2, dbo.Notes.Electricity,"
MySQL = MySQL & "                       dbo.Notes.BankID, dbo.BanksData.BankNamee, dbo.BanksData.BankName AS BankName2, dbo.TblNotesSales.rate, dbo.TblNotesSales.valu,"
MySQL = MySQL & "                       dbo.TblNotesSales.Type, dbo.TblNotesSales.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.Servce,"
 MySQL = MySQL & "                      dbo.Notes.RemaiValue, dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.RentValuePayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.CommissionsPayed, dbo.ContracttBillInstallmentsDone.InsurancePayed, dbo.ContracttBillInstallmentsDone.ElectricPayed,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.TelandNetPayed, dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.total, dbo.ContracttBillInstallmentsDone.[Value], dbo.ContracttBillInstallmentsDone.InstallNo,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.VATPayed, dbo.ContracttBillInstallmentsDone.VATValue, dbo.ContracttBillInstallmentsDone.ActVAT,"
MySQL = MySQL & "                       dbo.ContracttBillInstallmentsDone.Commisionvalue , dbo.ContracttBillInstallmentsDone.OldValuePayed, dbo.ContracttBillInstallmentsDone.PaymentType"
MySQL = MySQL & " FROM         dbo.ContracttBillInstallmentsDone RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.Notes ON dbo.ContracttBillInstallmentsDone.NoteID = dbo.Notes.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblNotesSales LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblNotesSales.EmpID = dbo.TblEmployee.Emp_ID ON dbo.Notes.NoteID = dbo.TblNotesSales.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_1 ON dbo.Notes.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_1 ON dbo.Notes.akarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqarDetai LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAkarUnit TblAkarUnit_2 ON dbo.TblAqarDetai.unittype = TblAkarUnit_2.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAqar TblAqar_2 ON dbo.TblAqarDetai.Aqarid = TblAqar_2.Aqarid ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "                                     LEFT OUTER JOIN dbo.TblUsers AS tu"
MySQL = MySQL & "                                   ON  dbo.Notes.UserID = tu.UserID"
'Where (dbo.Notes.NoteID = 4441)
MySQL = MySQL & " Where "
MySQL = MySQL & " (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "        AND ISNULL(contNo, 0) <> 0"

db_createOrUpdateviewSQL "View_Waiver", MySQL

  If DB_CreateTable("TblUsersProductLine", True, "id", True) = True Then
            DB_CreateField "TblUsersProductLine", "ProductLineId", adInteger, adColNullable, , , "  ", False, True
            DB_CreateField "TblUsersProductLine", "userid", adInteger, adColNullable, , , "  ", False, True
        End If

        DB_CreateField "TblProductLine", "FormPrint", adInteger, adColNullable, , , "", False, True
        
DB_CreateField "TblOptions", "PrintInvoiceByBranch", adInteger, adColNullable, , , "  ", False, True

 
DB_CreateField "TblOptions", "DefaultIsCreditPurchaseRet", adBoolean, adColNullable, , , " «·«› —«÷Ì «·»Ì⁄ «Ã·          ", False, True
    
  '»Œ’Ê’ „Ê÷Ê⁄  «—ÌŒ «·’·«ÕÌ…
    DB_CreateField "tblitems", "EXpirType", adInteger, adColNullable, , , "        ", False, True
    DB_CreateField "tblitems", "EXpireValue", adInteger, adColNullable, , , "        ", False, True
DB_CreateField "TblUsers", "StoreID2", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblUsers", "StoreID3", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblUsers", "StoreID2", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblReCostCalc", "EntryCreated", adBoolean, adColNullable, , , "      ", False, True


DB_CreateField "TblEmployee", "PrefNatID", adVarWChar, adColNullable, 400, , "      ", False, True, , True

DB_CreateField "TblFiterWaiver", "LastInvoiceRead", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "LastInvoiceRead2", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "Diff", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "Price", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "R", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "PrevBalance", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "ServiceCounter", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "TotalCounter", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "Notes", "IqarID2", adInteger, adColNullable, , , "  ", False, True
  DB_CreateField "notes_all", "IqarID2", adInteger, adColNullable, , , "  ", False, True
  
DB_CreateField "TblItemsParts", "Calories", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItems", "TotalCalories", adDouble, adColNullable, , , "    ", False, True

  DB_CreateField "TblItemsParts", "Calories", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItems", "TotalCalories", adDouble, adColNullable, , , "    ", False, True


    DB_CreateField "TblDefComItem", "Period", adDouble, adColNullable, , , "    ", False, True


 add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "44,'„’—Ê›«  «·„”«Â„« ','Trading Contract'", "ID", 44
 add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "45,'„—œÊœ«  „’—Ê›«  «·„”«Â„« ','Trading Contract'", "ID", 45
 add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "46,'›« Ê—…«·„»Ì⁄«  „”«Â„« ','Trading Contract'", "ID", 46
 DB_CreateField "TblExpensesInvesment", "VATyo", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "TblExpensesInvesment", "AccountCodeVat", adVarWChar, adColNullable, 55, , "      ", False, True, , True
 DB_CreateField "TblExpensesInvesmentDet", "VAT", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "TblExpensesInvesmentDet", "Total", adDouble, adColNullable, , , "    ", False, True

If DB_CreateTable("TblItemProductLine", True, "ID", True) = True Then
        DB_CreateField "TblItemProductLine", "ItemID", adInteger, adColNullable, , , " ???   ", False, True
        DB_CreateField "TblItemProductLine", "ProductLineId", adInteger, adColNullable, 10, , "C?C??   ", False, True, , True
       
        DB_CreateField "TblItemProductLine", "Remarks", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
         
End If

        DB_CreateField "TblProductLineDistribution", "BaseProductLineID", adInteger, adColNullable, , , "  ", False, True

    

  If DB_CreateTable("TblProductLineDistributionDet", True, "ID ", False) = True Then
        DB_CreateField "TblProductLineDistributionDet", "IDDefCIT", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistributionDet", "ProductLineID", adInteger, adColNullable, , , "  ", False, True
         DB_CreateField "TblProductLineDistributionDet", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistributionDet", "BaseProductLineID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistributionDet", "Qty", adDouble, adColNullable, , , "    ", False, True
End If

 
DB_CreateField "TblEmployee", "DOB", adDBTimeStamp, adColNullable, , , "      ", False, True
 DB_CreateField "TblMaintenanceWork", "DeptID", adInteger, adColNullable, , , "  ", False, True
 
 
   DB_CreateField "TblOptions", "EmpProduction", adBoolean, adColNullable, , , "", False, True
  DB_CreateField "TblOptions", "ItemProduction", adBoolean, adColNullable, , , "", False, True
  DB_CreateField "TblOptions", "ExpProduction", adBoolean, adColNullable, , , "", False, True
  

        DB_CreateField "Transactions", "CostForProductionItem", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "Transactions", "CostForProductionEmp", adDouble, adColNullable, , , "", False, True
    DB_CreateField "Transactions", "CostForProductionExp", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "Transactions", "CostForProductionTotal", adDouble, adColNullable, , , " ", False, True

 DB_CreateField "Transactions", "CusBalance", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
 
DB_CreateField "TblUsers", "CaNUpdateApprovedDoc", adBoolean, adColNullable, , , "", False, True




DB_CreateField "TblContract", "AccountCodeVat2", adVarWChar, adColNullable, 55, , "C?C??   ", False, True, , True
 DB_CreateField "TblContract", "FATYou2", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "TblContract", "Remark2", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
 
 
 
DB_CreateField "TblDefComItem", "BuiltinItemID", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblDefComItem", "GroupIDBuiltin", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblDefComItemData", "GroupIDBuiltin", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblDefComItemData", "BuiltinItemID", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblFiterWaiverDet2", "TypeDate", adInteger, adColNullable, , , " ???   ", False, True


 
DB_CreateField "TblFiterWaiver", "TotalinsuranceS", adDouble, adColNullable, , , " ???   ", False, True
DB_CreateField "TblFiterWaiver", "TypeMonthCalc", adBoolean, adColNullable, , , " ???   ", False, True


 
DB_CreateField "Transactions", "BankID", adInteger, adColNullable, , , " ???   ", False, True
 DB_CreateField "TblFiterWaiver", "TotalInsurances", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "Notes", "TotalInsurances", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "TblCustemers", "BrithDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
 DB_CreateField "TblCustemers", "BrithDate", adDBTimeStamp, adColNullable, , , "      ", False, True
     sql = "  DROP FUNCTION GetCountAllUnit" & CHR(13)
    Cn.Execute sql
    sql = " CREATE FUNCTION GetCountAllUnit(@ID integer  )"
    sql = sql & "  RETURNS Float"
    sql = sql & " AS"
    sql = sql & " Begin"
    sql = sql & " RETURN ( SELECT     COUNT(Id) AS CounNo"
    sql = sql & "  From dbo.TblAqarDetai"
    sql = sql & " WHERE      (Aqarid = @ID)"
    sql = sql & "   )"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "GetCountAllUnit", sql

add_record_to_table "TransactionTypes", "Transaction_Type,TransactionTypeName,TransactionEnglishName,StockEffect", " 75 , '  ”⁄Ì— «·«‰ «Ã ' , 'Production Pricing ' ,0", "Transaction_Type", 75


        DB_CreateField "TblDefComItem", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True
DB_CreateField "TblDefComItem", "RecDate", adDBTimeStamp, adColNullable, , , "      ", False, True

 DB_CreateField "Notes", "DebitSide", adVarWChar, adColNullable, 50, , "C?C??   ", False, True, , True
  DB_CreateField "Notes", "CreditSide", adVarWChar, adColNullable, 50, , "C?C??   ", False, True, , True
  
 
DB_CreateField "TblDefComItem", "TransactionID4", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblDefComItem", "NoteSerial14", adVarWChar, adColNullable, 255, , "      ", False, True, , True

DB_CreateField "Transaction_Details", "ProjectID", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "TblDefComItemData", "[Length]", adDouble, adColNullable, , , , False, True
DB_CreateField "Transaction_Details", "[Length]", adDouble, adColNullable, , , , False, True
DB_CreateField "TblDefComItem", "[Length]", adDouble, adColNullable, , , , False, True



DB_CreateField "TblFiterWaiver", "VAtPercent", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "VAt2", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblFiterWaiver", "TotalCounterNet", adDouble, adColNullable, , , "    ", False, True



 


DB_CreateField "TblDefComItemData", "TransactionID4", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblDefComItemData", "NoteSerial14", adVarWChar, adColNullable, 255, , "      ", False, True, , True


 


DB_CreateField "Transaction_Details", "ItemID2", adInteger, adColNullable, , , "    ", False, True



 DB_CreateField "TblDefComItemData", "PercentCost", adDouble, adColNullable, , , " ???    ", False, True
DB_CreateField "Transaction_Details", "PercentCost", adDouble, adColNullable, , , " ???    ", False, True
DB_CreateField "Transaction_Details", "GroupID", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "Transaction_Details", "TransactionID4", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "Transaction_Details", "NoteSerial14", adVarWChar, adColNullable, 50, , "  ", False, True, , True
DB_CreateField "Transaction_Details", "ItemID2", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "ContractVaueHousing", adDouble, adColNullable, , , " ???    ", False, True



DB_CreateField "TblDefComItem", "TransactionID5", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblDefComItem", "NoteSerial15", adVarWChar, adColNullable, 50, , "  ", False, True, , True
DB_CreateField "TblOptions", "OpenAccountAqar", adBoolean, adColNullable, , , " › Õ Õ”«» ·ﬂ· ⁄ﬁ«—  ", False, True
DB_CreateField "TblCustemers", "AccountAccountAqar", adVarWChar, adColNullable, 50, , "C?C??   ", False, True, , True
DB_CreateField "TblAqar", "AccounCode", adVarWChar, adColNullable, 50, , "C?C??   ", False, True, , True


            DB_CreateField "TblItemsParts", "Increase", adDouble, adColNullable, , , "    ", False, True
            
        DB_CreateField "TblDefComItemDet", "lowering", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblDefComItemDet", "Increase", adDouble, adColNullable, , , "    ", False, True
            
            




            
        DB_CreateField "TblDefComItemData", "Diameter", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "thickness", adDouble, adColNullable, , , "    ", False, True



       
        DB_CreateField "TblDefComItemData", "Diameter2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "thickness2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "widtj2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "DO2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "DI2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "hight2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblDefComItemData", "Length2", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblAdditionsAssest", "NoteSerial1", adVarWChar, adColNullable, 255, , "", False, True, , True

            
   
DB_CreateField "ContainerContracts", "RecType", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "ContainerContracts", "Status", adInteger, adColNullable, , , " ???    ", False, True

        DB_CreateField "ContainerContracts", "Contract_period_no", adInteger, adColNullable, , , "", False, True
        DB_CreateField "ContainerContracts", "Contract_period", adInteger, adColNullable, , , "", False, True
        DB_CreateField "ContainerContracts", "StrDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "ContainerContracts", "EndDate", adDBTimeStamp, adColNullable, 8, , "", False
 


 DB_CreateField "ContainerContractsDet", "NoDays", adDouble, adColNullable, , , " ???    ", False, True
 DB_CreateField "ContainerContractsRecDet", "NoDays", adDouble, adColNullable, , , " ???    ", False, True


   DB_CreateField "ContainerContractsRecDet", "RepliesNoFree", adVarWChar, adColNullable, 50, , " ???    ", False, True
       DB_CreateField "ContainerContractsDet", "RepliesNoFree", adVarWChar, adColNullable, 50, , " ???    ", False, True
       DB_CreateField "ContainerContractsRecDet", "RepliesValue", adDouble, adColNullable, 50, , " ???    ", False, True
       DB_CreateField "ContainerContractsDet", "RepliesValue", adDouble, adColNullable, 50, , " ???    ", False, True
       




DB_CreateField "TblAdditionsAssest", "NoteSerial1", adVarWChar, adColNullable, 255, , "", False, True, , True

            
   DB_CreateField "Notes", "TotalPayedOpBalance", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "TotalPayments", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "Percentage", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "NetValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "PreBalaValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "PreBalaPayed", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "PreBalaRemain", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "PreBalaTransPyed", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "PreBalaNet", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "TblExpensesDet", "TotalPayed", adInteger, adColNullable, , , "  ", False, True
 
 
DB_CreateField "Tbl_TradingContract", "NewMeasureNo", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "Tbl_TransOrder", "TradingContractID", adInteger, adColNullable, , , " ???    ", False, True


   
    DB_CreateField "TblFiterWaiver", "LastInstalldate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblFiterWaiver", "InstalldateH", adVarWChar, adColNullable, 10, , "C?C??   ", False, True, , True
    DB_CreateField "TblFiterWaiver", "ComResid", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblFiterWaiver", "TypeDate", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblFiterWaiver", "CalcLastPayment", adBoolean, adColNullable, , , " ??     ???? E??C? C?C??C? ", False, True
    



DB_CreateField "Tbl_TradingContract", "NewMeasureNo", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "Tbl_TransOrder", "TradingContractID", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblItems", "increase", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblItems", "lowering", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "Tbl_TradingContract", "PercentAlarm", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblUsers", "CaNUpdateAutoSalesInvoice", adBoolean, adColNullable, , , "", False, True
'

Cn.Execute " ALTER TABLE TblPrintBarCode DROP  column ProductionDate"
DB_CreateField "TblPrintBarCode", "ProductionDate", adDBTimeStamp, adColNullable, , , "", False, True

 

DB_CreateField "Transactions", "Chasee", adVarWChar, adColNullable, 10, , "C?C??   ", False, True, , True
 DB_CreateField "Transactions", "KM", adInteger, adColNullable, , , " ???    ", False, True

update_record_to_table "TblNotesTypes", "NotesTypeNamee", "Receipts", "NotesType", 4

DB_CreateField "TblDefComItem", "Noteid3", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "TblOptions", "InvoiceTransferJLTotal", adBoolean, adColNullable, , , "    «·ﬁÌœ «Ã„«·Ì ›Ì ›Ê« Ì— ⁄„·«¡ «·‰ﬁ·Ì«   ", False, True




 
        If DB_CreateTable("Tbl_TradingContractInv", True, "ID", False) = True Then
            DB_CreateField "Tbl_TradingContractInv", "TradingContractID", adInteger, adColNullable, , , " ??? ", False, True
            DB_CreateField "Tbl_TradingContractInv", "BD_Date", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "Tbl_TradingContractInv", "BD_BranchID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInv", "UserID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInv", "BD_Notes", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "Tbl_TradingContractInv", "BD_ProjectID", adInteger, adColNullable, , , " ???    ", False, True
                    DB_CreateField "Tbl_TradingContractInv", "NoteID", adInteger, adColNullable, , , "      ", False, True, , True


                    
        End If

        If DB_CreateTable("Tbl_TradingContractInvDet", True, "ID", True) = True Then
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_BD_ID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_BandNo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Qun", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Total", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Name", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_NameE", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Price", adDouble, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Vatyo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Vat2", adDouble, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "TotalNet", adDouble, adColNullable, , , " ???    ", False, True
            
                        DB_CreateField "Tbl_TradingContractInvDet", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
            
            DB_CreateField "Tbl_TradingContractInvDet", "TypeVAT", adVarWChar, adColNullable, 1000, , "      ", False, True, , True
            
            
        End If
        
        DB_CreateField "Tbl_TradingContractInvDet", "TConID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_BandNo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Qun", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Total", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Name", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_NameE", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "Tbl_TradingContractInvDet", "BDet_Price", adDouble, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Vatyo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "Vat2", adDouble, adColNullable, , , " ???    ", False, True
            DB_CreateField "Tbl_TradingContractInvDet", "TotalNet", adDouble, adColNullable, , , " ???    ", False, True
            
                        DB_CreateField "Tbl_TradingContractInvDet", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
                        
 
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3335 ,' ›Ê« Ì— „ﬁ«Ì”«  ' ,'      Trading Contract Invoice' ", "NotesType", 3378



DB_CreateField "TblProjecInvestment", "Path_General_photo", adVarWChar, adColNullable, 4000, , "      ", False, True, , True


DB_CreateField "TblItems", "IsPriceIsLenthW", adBoolean, adColNullable, , , "  ", False, True



add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3378 ,'   ???? ?????????? ' ,' Trading Contract Invoice' ", "NotesType", 3378
DB_CreateField "Tbl_TradingContractInv", "RecordDate", adDBTimeStamp, adColNullable, , , " ", False, True
DB_CreateField "Tbl_TradingContractInv", "NoteSerial1", adVarWChar, adColNullable, 255, , "      ", False, True, , True



DB_CreateField "Tbl_TradingContractInvDet", "SalPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "InstallPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "TotalSalPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "TotalInstallPrice", adDouble, adColNullable, , , "", False, True

    DB_CreateField "Transactions", "ProductionOrderID", adInteger, adColNullable, , , "      ", False, True

add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3378 ,'   ???? ?????????? ' ,' Trading Contract Invoice' ", "NotesType", 3378
DB_CreateField "Tbl_TradingContractInv", "RecordDate", adDBTimeStamp, adColNullable, , , " ", False, True
DB_CreateField "Tbl_TradingContractInv", "NoteSerial1", adVarWChar, adColNullable, 255, , "      ", False, True, , True



DB_CreateField "Tbl_TradingContractInvDet", "SalPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "InstallPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "TotalSalPrice", adDouble, adColNullable, , , "", False, True
DB_CreateField "Tbl_TradingContractInvDet", "TotalInstallPrice", adDouble, adColNullable, , , "", False, True

    DB_CreateField "Transactions", "ProductionOrderID", adInteger, adColNullable, , , "      ", False, True





   If DB_CreateTable("ShiftMaintType", True, "ID", False) = True Then
           DB_CreateField "ShiftMaintType", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "ShiftMaintType", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If


                   If DB_CreateTable("ShiftRec", True, "ID", False) = True Then
          DB_CreateField "ShiftRec", "OrderMaintinNo", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ShiftRec", "CustTel", adVarWChar, adColNullable, 50, , " ???    ", False, True
          DB_CreateField "ShiftRec", "ShiftMaintTypeID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ShiftRec", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
          DB_CreateField "ShiftRec", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ShiftRec", "typemaint", adInteger, adColNullable, , , " ???    ", False, True
                      DB_CreateField "ShiftRec", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
          
          DB_CreateField "ShiftRec", "DateRec", adDBTimeStamp, adColNullable, , , "      ", False, True
          

        DB_CreateField "ShiftRec", "TimeRec", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        
          DB_CreateField "ShiftRec", "Remarks", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ShiftRec", "NoteDone", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ShiftRec", "NoteStill", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ShiftRec", "NoteLate", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ShiftRec", "UserID", adInteger, adColNullable, , , " ???    ", False, True
       End If
       
       
       
       add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "47,'ÿ·» ’—› „ ⁄ÂœÌ‰','Request for disbursement of contractors'", "ID", 47


Cn.Execute " UPDATE VatTypes SET VatTypeName = ' «” Õﬁ«ﬁ«  «·„ ⁄ÂœÌ‰' WHERE ID = 2"


DB_CreateField "TblExchangeReques_Detailst", "TypeVAT", adVarWChar, adColNullable, 400, , "    ", False, True
DB_CreateField "TblExchangeReques_Detailst", "Vatyo", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblExchangeReques_Detailst", "TotalNet", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblExchangeReques_Detailst", "Vat2", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblAttributionInstallmentDivided", "TypeVAT", adVarWChar, adColNullable, 400, , "    ", False, True
DB_CreateField "TblAttributionInstallmentDivided", "Vatyo", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAttributionInstallmentDivided", "TotalNet", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAttributionInstallmentDivided", "Vat2", adDouble, adColNullable, , , "    ", False, True


  add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", "-1 ,'     ’›Ì… «·⁄ﬁ«—     ' ,'  ’›Ì… «·⁄ﬁ«— ' ", "NotesType", -1
  DB_CreateField "Notes", "RemainRent", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "RemainWater", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "BillPrice", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "RemainCommissions", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "OldRent", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "RemainService", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "insurance", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "txtOldInsurance", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Notes", "FilterID2", adInteger, adColNullable, , , "  ", False, True
 
 DB_CreateField "TblNotesOwnerPayment202", "NoteType", adInteger, adColNullable, , , "  ", False, True
 DB_CreateField "TblOwnerPayment202", "NoteType", adInteger, adColNullable, , , "  ", False, True
  DB_CreateField "TblDefComItem", "MaxNo2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
                                   
DB_CreateField "tblGeneralCashing", "CashierId", adInteger, adColNullable, , , "  ", False, True
 
DB_CreateField "TblOptions", "GeneralVoucherCreateSalesGE", adInteger, adColNullable, , , "  ", False, True



DB_CreateField "TblRevenuesTypes", "ManualEntrty", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "notes_all", "chkvat", adInteger, adColNullable, , , " ???    ", False, True

                 
                           If DB_CreateTable("TblCardAuthorizationReformItems", True, "ID2", True) = True Then
                            DB_CreateField "TblCardAuthorizationReformItems", "ID", adInteger, adColNullable, , , ""
                            DB_CreateField "TblCardAuthorizationReformItems", "ItemID", adInteger, adColNullable, , , ""
                            DB_CreateField "TblCardAuthorizationReformItems", "Vatyo", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "Price", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "Vat2", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "TotalWithVat", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "Remark", adVarWChar, adColNullable, 4000, , "??C?UE ?EI??E    ", False
                         End If
                         DB_CreateField "TblCardAuthorizationReformItems", "qty", adDouble, adColNullable, , , "    ", False, True
                         DB_CreateField "TblCardAuthorizationReformItems", "beforeVat", adDouble, adColNullable, , , "    ", False, True
                         
                       
                          DB_updateField "Transactions", "Chasee", "nvarchar(4000)   "
                       '  DB_updateField "Transaction_Details", "Chasee", "nvarchar(4000)   "
                              DB_CreateField "TblUsers", "CanChangeStatusDateRequest", adBoolean, adColNullable, , , "", False, True
                              

    sql = "    DROP FUNCTION QryItemsInventry3" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION [dbo].[QryItemsInventry3] (@fromdate datetime,@todate datetime,@StoreId AS INT=null,@ColorID AS INT=null,@ItemSize AS  NVARCHAR(255)=null ,"
    sql = sql & "  @ClassId AS INT=null , @order_no  AS  NVARCHAR(255)  =null,@CusID as float=null)"
    sql = sql & " RETURNS @XTable Table"
    sql = sql & "    ("
    sql = sql & "      Item_ID  Decimal (18,2),"
    sql = sql & "   LotNO  nvarchar(255),"
    sql = sql & " ItemCode     nvarchar(255)  ,"
    sql = sql & "     ItemName  nvarchar(4000)     ,"
    sql = sql & "   openingValue  Decimal (18,2),"
    sql = sql & "    inputvalue  Decimal (18,2),"
    sql = sql & "   outputValue Decimal(18, 2)"
    sql = sql & " )"
    sql = sql & "  AS"
    sql = sql & " Begin"
    sql = sql & " INSERT  @XTable"
    sql = sql & " Select      Item_ID,  LotNO, ItemCode, ItemName,  Sum(DEV_Value1) as openingValue,Sum(DEV_Value2) as inputvalue , Sum(DEV_Value3) as outputValue"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & " SELECT"
    sql = sql & " Item_ID,  null as LotNO, ItemCode, ItemName,"
    sql = sql & " DEV_Value1=Case"
    sql = sql & "  When  (dbo.TransactionTypes.StockEffect = 1)  and (dbo.Transactions.Transaction_Type=3)   Then  (Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & "  END,"
    sql = sql & "  DEV_Value2=Case"
    sql = sql & "  When  (dbo.TransactionTypes.StockEffect = 1)  and (dbo.Transactions.Transaction_Type<>3)   Then  (Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & "  End"
    sql = sql & "   ,"
    sql = sql & " DEV_Value3=Case"
    sql = sql & " When    (dbo.TransactionTypes.StockEffect = -1)   Then ( Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & " End"
    sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    sql = sql & " INNER JOIN  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & " INNER JOIN    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    sql = sql & " where dbo.Transactions.Transaction_Date>=@fromdate"
    sql = sql & " and dbo.Transactions.Transaction_Date <=@todate"
    sql = sql & " and  Storeid=isnull(@Storeid,Storeid)"
    sql = sql & " and  ColorID=isnull(@ColorID,ColorID)"
    sql = sql & " and  ItemSize=isnull(@ItemSize,ItemSize)"
    sql = sql & " and  ClassId=isnull(@ClassId,ClassId)"
    sql = sql & " )XTable"
    sql = sql & " group by Item_ID ,LotNO,ItemCode, ItemName"
    sql = sql & " Return"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "QryItemsInventry3", sql
                              
       MySQL = " SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
 MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
 MySQL = MySQL & "                      dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType,"
 MySQL = MySQL & "                      dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Parent_Account_Code,"
 MySQL = MySQL & "                      dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center,"
 MySQL = MySQL & "                      dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
 MySQL = MySQL & "                      dbo.TblNotesTypes.NotesTypeNamee, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
 MySQL = MySQL & "                      dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo"
 MySQL = MySQL & "    FROM         dbo.TblAqar RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBranchesData INNER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers INNER JOIN"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON"
 MySQL = MySQL & "                      dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.Notes LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
 MySQL = MySQL & "     Where (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
  db_createOrUpdateviewSQL "RptLedger_Sub", MySQL
  
  
 DB_CreateField "ContainerContracts", "RecDate", adDBTimeStamp, adColNullable, , , "      ", False, True
   DB_CreateField "ShiftRec", "TimeEnd", adVarWChar, adColNullable, 50, , "      ", False, True, , True
               DB_CreateField "ShiftRec", "DateEnd", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "ShiftRec", "CarStatus", adInteger, adColNullable, , , " ???    ", False, True





    DB_CreateField "Transactions", "RecConditions", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "Transactions", "DownloadPort", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "Transactions", "PortLoading", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "Transactions", "PaymentConditions", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "Transactions", "ShipName", adVarWChar, adColNullable, 400, , "      ", False, True, , True


    DB_CreateField "Transactions", "PaymentConditions", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Transactions", "DeliveryMethod", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "Transactions", "Packing", adVarWChar, adColNullable, 400, , "      ", False, True, , True


                DB_CreateField "ShiftRec", "TimeEnd", adVarWChar, adColNullable, 50, , "      ", False, True, , True
               DB_CreateField "ShiftRec", "DateEnd", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "ShiftRec", "CarStatus", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblEmployee", "Account_codeTEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblEmployee", "Account_code1TEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblEmployee", "Account_code2TEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblEmployee", "Account_code3TEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblEmployee", "Account_code4TEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblEmployee", "Account_code5TEMP", adVarWChar, adColNullable, 255, , "      ", False, True, , True



add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "48,'«·Õ«ÊÌ« ','Containers'", "ID", 48


add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3379 ,'«·Õ«ÊÌ«     ' ,'      Containers' ", "NotesType", 3379



add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "49,'«· ›—Ì€« ','Containers Unloading'", "ID", 49


      DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "ContainerContracts", "NoteID", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "ContainerContracts", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True

        DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "branches", "a209", adVarWChar, adColNullable, 50, , "", False, True, , True
DB_CreateField "branches", "a210", adVarWChar, adColNullable, 50, , "", False, True, , True



DB_CreateField "branches", "a211", adVarWChar, adColNullable, 250, , "", False, True, , True
DB_CreateField "branches", "a212", adVarWChar, adColNullable, 250, , "", False, True, , True

add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3380 ,'«· ›—Ì€«     ' ,'      Containers' ", "NotesType", 3380


'add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "48,'«·Õ«ÊÌ« ','Containers'", "ID", 48


    If DB_CreateTable("ContainerUnloading", True, "ID", False) = True Then
          DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "CustTel", adVarWChar, adColNullable, 50, , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "ContractNo", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
          DB_CreateField "ContainerUnloading", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Remarks", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ContainerUnloading", "UserID", adInteger, adColNullable, , , " ???    ", False, True
          
          DB_CreateField "ContainerUnloading", "Value", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Count", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscPercent", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "NetBDisc", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Vat", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalNet", adDouble, adColNullable, , , " ???    ", False, True
          
        DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True
       End If

  


add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "48,'«·Õ«ÊÌ« ','Containers'", "ID", 48
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3379 ,'«·Õ«ÊÌ«     ' ,'      Containers' ", "NotesType", 3379
add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "49,'«· ›—Ì€« ','Containers Unloading'", "ID", 49
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3380 ,'«· ›—Ì€«     ' ,'      Containers' ", "NotesType", 3380
      
DB_CreateField "branches", "a209", adVarWChar, adColNullable, 50, , "", False, True, , True
DB_CreateField "branches", "a210", adVarWChar, adColNullable, 50, , "", False, True, , True
      DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "ContainerContracts", "RecDate", adDBTimeStamp, adColNullable, , , "      ", False, True

DB_CreateField "ContainerContracts", "NoteID", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "ContainerContracts", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True

        DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True
                 DB_CreateField "ContainerContracts", "Blanks", adDouble, adColNullable, , , "    ", False, True









DB_CreateField "ContainerContractsRecDet", "GroupID2", adInteger, adColNullable, , , " ???    ", False, True
    'If DB_CreateTable("ContainerUnloading", True, "ID", False) = True Then
          DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "CustTel", adVarWChar, adColNullable, 50, , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "ContractNo", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
          DB_CreateField "ContainerUnloading", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Remarks", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ContainerUnloading", "UserID", adInteger, adColNullable, , , " ???    ", False, True
          
          DB_CreateField "ContainerUnloading", "Value", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Count", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscPercent", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "NetBDisc", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Vat", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalNet", adDouble, adColNullable, , , " ???    ", False, True
          
        DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    '   End If

  

  
DB_CreateField "TblUsers", "CanChangeTripAfterInvoiceing", adBoolean, adColNullable, , , "", False, True


                            DB_CreateField "TblCardAuthorizationReformItems", "PriceBDisc", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "DiscPercent", adDouble, adColNullable, , , "    ", False, True
                            DB_CreateField "TblCardAuthorizationReformItems", "DiscValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblCardAuthorizationReformItems", "DiscValue", adDouble, adColNullable, , , "    ", False, True


    If DB_CreateTable("tblCustomerFingers", True, "ID", True) = True Then
          DB_CreateField "tblCustomerFingers", "FCusID", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "tblCustomerFingers", "ItemPhoto", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True
DB_CreateField "tblCustomerFingers", "ItemPhoto1", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True
DB_CreateField "tblCustomerFingers", "ItemPhoto2", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True
DB_CreateField "tblCustomerFingers", "ItemPhoto3", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True


 End If
 
 DB_CreateField "TblCardAuthorizationReformItems", "ItemName2", adVarWChar, adColNullable, 4000, , "??C?UE ?EI??E    ", False




             DB_CreateField "TblCardAuthorizationReform", "DiscValue", adDouble, adColNullable, , , "    ", False, True
                        DB_CreateField "TblCardAuthorizationReform", "DiscPercent", adDouble, adColNullable, , , "    ", False, True
                        DB_CreateField "TblCardAuthorizationReform", "TotalAfterDiscount", adDouble, adColNullable, , , "    ", False, True
                        DB_CreateField "TblCardAuthorizationReform", "Vatyo", adDouble, adColNullable, , , "    ", False, True
                        DB_CreateField "TblCardAuthorizationReform", "Vat2", adDouble, adColNullable, , , "    ", False, True
                        DB_CreateField "TblCardAuthorizationReform", "Vat2", adDouble, adColNullable, , , "    ", False, True



DB_CreateField "TblFiterWaiver", "RenterID2", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "emp_salary", "VoCation2", adInteger, adColNullable, , , "   ", False, True
DB_CreateField "emp_salary", "VoCation4", adInteger, adColNullable, , , "   ", False, True
DB_CreateField "emp_salary", "VoCation3", adInteger, adColNullable, , , "   ", False, True
DB_CreateField "TblOptions", "SalesNotCreateGe", adInteger, adColNullable, , , "  ", False, True
 DB_updateField "notes_all", "RecNo", "nvarchar(4000)   "
  DB_CreateField "TblExchange", "BankIBAN", adVarWChar, adColNullable, 4000, , "??C?UE ?EI??E    ", False

DB_CreateField "Transaction_Details", "ShowAttatch", adDouble, adColNullable, , , "  ", False, True
DB_CreateField "TblUsers", "CanChangeOut", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblCarMaintenancePlanDetails", "CancelReason", adVarWChar, adColNullable, 4000, , "??C?UE ?EI??E    ", False

    DB_CreateField "TBLInsurancesJoin", "WorkDays", adDouble, adColNullable, , , , False, True
   DB_CreateField "TBLInsurancesJoin", "BignDateWork", adDBTimeStamp, adColNullable, , , "      ", False, True

DB_CreateField "ShiftRec", "StutsMaint", adInteger, adColNullable, , , "  ", False, True



 sql = "    DROP FUNCTION GetBalanceByproject" & CHR(13)
    Cn.Execute sql
   sql = " CREATE FUNCTION GetBalanceByproject(@fromdate datetime  ,@Todate datetime  ,@project_id as integer,@Account_Code as varchar(255) =null )"
   sql = sql & "  RETURNS Float" & CHR(13)
   sql = sql & " AS" & CHR(13)
   sql = sql & " Begin" & CHR(13)
   sql = sql & " RETURN (" & CHR(13)
   sql = sql & "  SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
   sql = sql & " FROM (" & CHR(13)
   sql = sql & " SELECT" & CHR(13)
   sql = sql & "  DEV_Value1=Case" & CHR(13)
   sql = sql & " When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
   sql = sql & " Else 0" & CHR(13)
   sql = sql & "  END," & CHR(13)
   sql = sql & "  DEV_Value2=Case" & CHR(13)
   sql = sql & "  When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
   sql = sql & "  Else 0" & CHR(13)
   sql = sql & " End" & CHR(13)
   sql = sql & " from dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN" & CHR(13)
   sql = sql & " dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
   sql = sql & "  WHERE    ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
   sql = sql & "  and dbo.DOUBLE_ENTREY_VOUCHERS.project_id =@project_id and  dbo.DOUBLE_ENTREY_VOUCHERS. Posted is null" & CHR(13)
  sql = sql & "    and  DOUBLE_ENTREY_VOUCHERS.Account_Code= isnull(@Account_Code,DOUBLE_ENTREY_VOUCHERS.Account_Code)"
    sql = sql & " )XTABLE" & CHR(13)
   sql = sql & "  )" & CHR(13)
   sql = sql & "  End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetBalanceByproject", sql
   
   DB_CreateField "Notes", "ContainerNo", adInteger, adColNullable, , , , False, True

DB_CreateField "LogFile", "ConnectionData", adVarWChar, adColNullable, 4000, , "??C?UE ?EI??E    ", False

DB_CreateField "TblOrderUpload", "TypeRep", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TblEmpPassOver", "GroupID", adInteger, adColNullable, , , "  ", False, True


DB_CreateField "TblChangedComponentRegister", "Remarks", adVarWChar, adColNullable, 400, , "      ", False, True, , True

   
  DB_CreateField "TblFiterWaiver", "ApartmentID2", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblFiterWaiver", "BulidID2", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblFiterWaiver", "unittype2", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblPripaidExpenses", "NewOrOpeneing", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblPripaidExpensesDet", "NewOrOpeneing", adInteger, adColNullable, , , " ???    ", False, True

        DB_CreateField "TblPaytAmortizationDet", "TotalVal", adDouble, adColNullable, , , "    ", False, True
     DB_CreateField "TblPaytAmortizationDet", "IDD", adInteger, adColNullable, , , " ???    ", False, True
     DB_CreateField "TblPaytAmortizationDet", "ChID", adInteger, adColNullable, , , " ???    ", False, True
     DB_CreateField "TblPripaidExpChiled", "ProfExpID", adInteger, adColNullable, , , " ???    ", False, True
     
   
      If DB_CreateTable("GranteeType", True, "ID", False) = True Then
           DB_CreateField "GranteeType", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "GranteeType", "NameE", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "GranteeType", "Period", adInteger, adColNullable, , , " ???    ", False, True
     End If




If DB_CreateTable("ItemsGranteeType", True, "ID", True) = True Then
        DB_CreateField "ItemsGranteeType", "ItemID", adInteger, adColNullable, , , " ???   ", False, True
        DB_CreateField "ItemsGranteeType", "GranteeTypeID", adInteger, adColNullable, 10, , "C?C??   ", False, True, , True
            DB_CreateField "ItemsGranteeType", "Period", adInteger, adColNullable, 10, , "C?C??   ", False, True, , True
       
        DB_CreateField "ItemsGranteeType", "Remarks", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
         
End If

   
    DB_updateField "TOrder_Notes", "TOrder_Notes", "nvarchar(4000)   "
   
   DB_CreateField "Notes", "installIDCont", adDouble, adColNullable, , , "    ", False, True
   
   DB_CreateField "jopstatus", "StopAllocation", adBoolean, adColNullable, , , "", False, True
   

'StopAllocation
DB_CreateField "TblOptions", "RawMaterMix2", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "notes_all", "AccountPaym", adVarWChar, adColNullable, 55, , "      ", False, True, , True

 

                DB_CreateField "TblProductLine", "StoreID", adInteger, adColNullable, , , "", False, True

        DB_CreateField "TblOptions", "DontCreateOut", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "DontCreateOut2", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "InsertItemManualOut", adBoolean, adColNullable, , , " ", False, True

 

                DB_CreateField "Transactions", "IDDefCIT", adInteger, adColNullable, , , "", False, True

                DB_CreateField "TblProductLine", "StoreID", adInteger, adColNullable, , , "", False, True

        DB_CreateField "TblOptions", "DontCreateOut", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "DontCreateOut2", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "InsertItemManualOut", adBoolean, adColNullable, , , " ", False, True

DB_CreateField "TblDefComItemDet", "TableID", adInteger, adColNullable, , , "", False, True

                DB_CreateField "Transactions", "IDDefCIT", adInteger, adColNullable, , , "", False, True

                DB_CreateField "TblProductLine", "StoreID", adInteger, adColNullable, , , "", False, True

        DB_CreateField "TblOptions", "DontCreateOut", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "DontCreateOut2", adBoolean, adColNullable, , , " ", False, True
        DB_CreateField "TblOptions", "InsertItemManualOut", adBoolean, adColNullable, , , " ", False, True





DB_CreateField "TblDefComItemDet", "QtyOut", adDouble, adColNullable, , , "", False, True

DB_CreateField "TblDefComItemDet", "TableID", adInteger, adColNullable, , , "", False, True

DB_CreateField "Transaction_Details", "RemarksLine", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

DB_CreateField "Transaction_Details", "LineID", adInteger, adColNullable, , , "", False, True

DB_CreateField "TblUsers", "CanCancelContract", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblOptions", "CarsRevenuePerOwner", adBoolean, adColNullable, , , "    «·ﬁÌœ «Ã„«·Ì ›Ì ›Ê« Ì— ⁄„·«¡ «·‰ﬁ·Ì«   ", False, True

DB_CreateField "TblCarsData", "AccountPaym", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True

    
           
    If DB_CreateTable("TblEndDebtAgingInvDet2", True, "ID", True) = True Then
        DB_CreateField "TblEndDebtAgingInvDet2", "EndDebAgInvID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "IsHeaderRec", adBoolean, adColNullable, , , "                ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "CusID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "NoteSerial1", adVarWChar, adColNullable, 100, , "    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "NoteSerial", adVarWChar, adColNullable, 100, , "    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "Note_Value", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "PayedValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "value", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "NotesTypeName", adVarWChar, adColNullable, 4000, , "    ", False, True
        DB_CreateField "TblEndDebtAgingInvDet2", "remark", adVarWChar, adColNullable, 4000, , "    ", False, True
        
        DB_CreateField "TblEndDebtAgingInvDet2", "too", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
        DB_CreateField "TblEndDebtAgingInvDet2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
       
        
        DB_CreateField "TblEndDebtAgingInvDet2", "NetValue", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblEndDebtAgingInvDet2", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
        
    End If
    
        
DB_CreateField "TblOptions", "DontShowMoreDetailsCompItem", adBoolean, adColNullable, , , " ", False, True


Cn.Execute "ALTER TABLE TblCustemers DROP  COLUMN  PaymentType "

DB_CreateField "TblCustemers", "CPaymentType", adInteger, adColNullable, , , "        ", False, True
           
   DB_CreateField "TblBalanceSheetHeader", "DYear", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblBalanceSheetDetails", "DebitValue", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblBalanceSheetDetails", "CreditValue", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblBalanceSheetDetails", "AValue", adDouble, adColNullable, , , "    ", False, True
   
DB_CreateField "TblUsers", "CanCustomerandVendor", adBoolean, adColNullable, , , "", False, True
    

DB_CreateField "Transactions", "EmpOrderInstitute", adInteger, adColNullable, , , "        ", False, True

 DB_updateField "TblTravDueKDet", "RecNo", "nvarchar(4000)   "
''DB_updateField "ACCOUNTS", "Account_Name", "nvarchar(4000)   "
DB_CreateField "TblOptions", "traveDiscountFromCustomerDirect", adBoolean, adColNullable, , , " ", False, True

If DB_CreateTable("GroupsCustomers", True, "GroupID", True) = True Then
    DB_CreateField "GroupsCustomers", "GroupName", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "GroupsCustomers", "GroupNamee", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "GroupsCustomers", "ParentId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "GroupsCustomers", "GroupCode", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "GroupsCustomers", "Code", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "GroupsCustomers", "FullCode", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "GroupsCustomers", "LastGroup", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "GroupsCustomers", "prifix", adVarWChar, adColNullable, 255, , "", False, True, , True
End If



    DB_CreateField "TblCustemers", "ClassCustomersId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustemers", "GroupsCustomersId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustemers", "BranchName", adVarWChar, adColNullable, 255, , " ", False, True, , True


   If DB_CreateTable("ClassCustomers", True, "ID", False) = True Then
           DB_CreateField "ClassCustomers", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "ClassCustomers", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If

    DB_CreateField "Groups", "ClassId1", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Groups", "ClassId2", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Groups", "ClassId3", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Groups", "ClassId4", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Groups", "ClassId5", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Groups", "ClassId6", adInteger, adColNullable, , , "    ", False, True





DB_CreateField "TblTravDueK", "chkoWithoutVat", adBoolean, adColNullable, , , "", False, True
     
     If DB_CreateTable("TmpItemsQty", True, "ID", False) = True Then
              DB_CreateField "TmpItemsQty", "LineID", adInteger, adColNullable, , , " ???    ", False, True
              DB_CreateField "TmpItemsQty", "Qty", adDouble, adColNullable, , , " ???    ", False, True
    End If
              
              
     
    DB_CreateField "TblOptions", "IsCustSalesManCashRelated", adBoolean, adColNullable, , , " ", False, True

    



 DB_CreateField "TblCustomerContract", "Emp_ID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblCustomerContract", "CashCustomerName", adVarWChar, adColNullable, 255, , "      ", False


DB_CreateField "TblOptions", "IsCustSalesManCashRelated", adBoolean, adColNullable, , , " ", False, True

              DB_CreateField "TblContract", "InsurValueInVAT", adInteger, adColNullable, , , " ???    ", False, True
              
             
    DB_CreateField "TblCarsData", "ExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCarsData", "SensitiveWeightDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
 DB_CreateField "TblUsers", "CanEditCars", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblEmployee", "IssuingDriverCardDate", adDBTimeStamp, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "CardDriverExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
        
              
    DB_CreateField "TblCarsData", "ExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCarsData", "SensitiveWeightDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
 DB_CreateField "TblUsers", "CanEditCars", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblEmployee", "IssuingDriverCardDate", adDBTimeStamp, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "CardDriverExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
 
        
 
 DB_CreateField "TblCarsData", "ExpireDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
 DB_CreateField "TblCarsData", "SensitiveWeightDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
       
    DB_CreateField "TblCarsData", "ExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCarsData", "SensitiveWeightDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
 DB_CreateField "TblUsers", "CanEditCars", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblEmployee", "IssuingDriverCardDate", adDBTimeStamp, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "CardDriverExpireDate", adDBTimeStamp, adColNullable, , , "      ", False, True
 
 
 DB_CreateField "TblCarsData", "ExpireDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
 DB_CreateField "TblCarsData", "SensitiveWeightDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True


                 
DB_CreateField "TblEmployee", "IssuingDriverCardDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
DB_CreateField "TblEmployee", "CardDriverExpireDateH", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True

           sql = "  DROP FUNCTION GetQurrentQtyInStock" & CHR(13)
    Cn.Execute sql
    sql = " CREATE FUNCTION GetQurrentQtyInStock(@StoreID integer,@ItemID integer ,@FrmDate datetime,@TODate datetime )" & CHR(13)
    sql = sql & "  RETURNS Float" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " RETURN (SELECT     SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS QNty" & CHR(13)
    sql = sql & "      FROM         dbo.Transactions INNER JOIN" & CHR(13)
    sql = sql & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN" & CHR(13)
    sql = sql & "                  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type" & CHR(13)
    sql = sql & "            WHERE     (dbo.Transactions.Transaction_Date >= @FrmDate)and (dbo.Transactions.Transaction_Date <= @TODate) AND "
    sql = sql & "  (dbo.Transactions.StoreID = @StoreID) AND" & CHR(13)
    sql = sql & "                  (dbo.Transaction_Details.Item_ID = @ItemID) AND (dbo.TransactionTypes.StockEffect <> 0)" & CHR(13)
    sql = sql & "   )" & CHR(13)
    sql = sql & " End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetQurrentQtyInStock", sql


If DB_CreateTable("TblGroupItemProductLine", True, "ID", True) = True Then
        DB_CreateField "TblGroupItemProductLine", "GroupID", adInteger, adColNullable, , , " ???   ", False, True
        DB_CreateField "TblGroupItemProductLine", "ProductLineId", adInteger, adColNullable, 10, , "C?C??   ", False, True, , True
       
        DB_CreateField "TblGroupItemProductLine", "Remarks", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
         
End If

DB_CreateField "TblOptions", "showEmployeeAccountIntrip", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblOptions", "DUEDOCUMENTbyinstallDate", adBoolean, adColNullable, , , " ", False, True


DB_CreateField "TblOptions", "CanSkipPurchOrder", adBoolean, adColNullable, , , " ", False, True



DB_CreateField "notes_all", "NoR", adDouble, adColNullable, , , " ", False, True

DB_CreateField "TblUsers", "HidLowering", adBoolean, adColNullable, , , "    ", False, True

DB_CreateField "TblOptions", "CompilingBasedTable", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblOptions", "DontSaveInvoiceWithoutDocType", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblOptions", "DontDuplicateManulaNoInPurchase", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "Transactions", "CIBAN", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True

    DB_CreateField "TblOptions", "CanPartialpayment", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblTravDueKDet", "notesallid", adDouble, adColNullable, , , " ???    ", False, True
         DB_CreateField "Transaction_Details", "ReCostID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "Transaction_Details", "FlgReCost", adInteger, adColNullable, , , "  ", False, True
     
     DB_CreateField "Transaction_Details", "ReCostID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "Transaction_Details", "FlgReCost", adInteger, adColNullable, , , "  ", False, True
     
     DB_CreateField "Transaction_Details", "ReCostID", adInteger, adColNullable, , , "  ", False, True
           DB_CreateField "Transaction_Details", "FlgReCost", adInteger, adColNullable, , , "  ", False, True
     
     DB_CreateField "Transactions", "chkAutoDetect", adBoolean, adColNullable, , , "", False, True
     

DB_CreateField "TblOptions", "EndRentifPayed", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblUsersProductLine", "ShowAlarm", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblDefComItemData", "DO", adDouble, adColNullable, , , , False, True
DB_CreateField "TblDefComItemData", "DI", adDouble, adColNullable, , , , False, True
DB_CreateField "TblOptions", "cantCahngeAkarinExpenses", adBoolean, adColNullable, , , "", False, True


If DB_CreateTable("TblGroupItemProductLineUsersset", True, "ID", True) = True Then
         
        DB_CreateField "TblGroupItemProductLineUsersset", "Username", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        DB_CreateField "TblGroupItemProductLineUsersset", "USERDOMAIN", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        DB_CreateField "TblGroupItemProductLineUsersset", "MACAddress", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        DB_CreateField "TblGroupItemProductLineUsersset", "computername", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        
         
End If


DB_CreateField "Transactions", "ContainerNo", adInteger, adColNullable, , , "", False, True

DB_CreateField "TblUsersProductLine", "TypeLine", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TblEmployee", "To_Employee_name", adVarWChar, adColNullable, 4000

DB_CreateField "TblOptions", "EmployeeSalaryBYBranch", adBoolean, adColNullable, , , "", False, True
DB_CreateField "Notes", "CCOPTion", adInteger, adColNullable, , , "  ", False, True


 UpdateDataBasePart25
  UpdateDataBasePart26
  
UpdateDataBasePart28
UpdateDataBasePart30
'UpdateDataBasePart27




End Function
Function UpdateDataBasePart25()


On Error Resume Next
 Dim New_View  As String
Dim s  As String

     If DB_CreateTable("LegalIssuesData", True, "ID", False) = True Then
           
            DB_CreateField "LegalIssuesData", "IssuesNo", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "LegalIssuesData", "RecordDate", adDBTimeStamp, adColNullable, , , ", False, True"
            
           DB_CreateField "LegalIssuesData", "CustID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "LegalIssuesData", "LegalcourtsID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "LegalIssuesData", "LegalIssuesID", adInteger, adColNullable, , , " ???    ", False, True
            
            DB_CreateField "LegalIssuesData", "IssuesReason", adVarWChar, adColNullable, 4000, , "     ", False, True, , True
            DB_CreateField "LegalIssuesData", "IssuesDesc", adVarWChar, adColNullable, 4000, , "     ", False, True, , True
            DB_CreateField "LegalIssuesData", "RecordDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
                        DB_CreateField "LegalIssuesTrans", "SessionDateNo", adVarWChar, adColNullable, 255, , "      ", False
     End If
     

      
    
    If DB_CreateTable("SessionDate", True, "ID", False) = True Then
           
            DB_CreateField "SessionDate", "IssuesNo", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "SessionDate", "RecordDate", adDBTimeStamp, adColNullable, , , ", False, True"
            DB_CreateField "SessionDate", "RecordDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            DB_CreateField "SessionDate", "SessionTime", adVarWChar, adColNullable, 50, , "C?C??   ", False, True, , True
            
            DB_CreateField "SessionDate", "SessionDate", adDBTimeStamp, adColNullable, , , ", False, True"
            DB_CreateField "SessionDate", "SessionDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
            
            
           DB_CreateField "SessionDate", "CustID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "SessionDate", "LegalcourtsID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "SessionDate", "LegalIssues", adInteger, adColNullable, , , " ???    ", False, True
            
            DB_CreateField "SessionDate", "SessionPlace", adVarWChar, adColNullable, 4000, , "     ", False, True, , True
            DB_CreateField "SessionDate", "IssuesDesc", adVarWChar, adColNullable, 4000, , "     ", False, True, , True

           
            
            
     End If


  If DB_CreateTable("LegalIssuesTrans", True, "ID", False) = True Then
            DB_CreateField "LegalIssuesTrans", "SessionDateNo", adVarWChar, adColNullable, 255, , "      ", False

            DB_CreateField "LegalIssuesTrans", "IssuesNo", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "LegalIssuesTrans", "RecordDate", adDBTimeStamp, adColNullable, , , ", False, True"
            
           DB_CreateField "LegalIssuesTrans", "CustID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "LegalIssuesTrans", "LegalcourtsID", adInteger, adColNullable, , , " ???    ", False, True
           DB_CreateField "LegalIssuesTrans", "LegalIssues", adInteger, adColNullable, , , " ???    ", False, True
            
            DB_CreateField "LegalIssuesTrans", "IssuesReason", adVarWChar, adColNullable, 4000, , "     ", False, True, , True
            DB_CreateField "LegalIssuesTrans", "SessionResult", adVarWChar, adColNullable, 4000, , "     ", False, True, , True
            DB_CreateField "LegalIssuesTrans", "RecordDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
            
     End If
     






   If DB_CreateTable("LegalIssues", True, "ID", False) = True Then
           DB_CreateField "LegalIssues", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "LegalIssues", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If

   If DB_CreateTable("Legalcourts", True, "ID", False) = True Then
           DB_CreateField "Legalcourts", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "Legalcourts", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If

Dim i As Integer
For i = 41 To 50
 add_record_to_table "Pmanger", "id", CStr(i), "id", i
 Next i
 

            DB_CreateField "Transactions", "PurchOrderNo", adVarWChar, adColNullable, 255, , "      ", False


  add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 9095 ,' „ «»⁄… «·—Õ·«     ' ,' Follow-up trips' ", "NotesType", 9095
    
        If DB_CreateTable("tblTripTrans", True, "ID", False) = True Then
            DB_CreateField "tblTripTrans", "recordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "tblTripTrans", "recordDateH", adVarWChar, adColNullable, 10, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans", "Fromdate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "tblTripTrans", "FromdateH", adVarWChar, adColNullable, 10, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans", "todate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "tblTripTrans", "todateH", adVarWChar, adColNullable, 10, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans", "BranchId", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans", "Remarks", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans", "BoxID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans", "PaymentType", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans", "NoteSerial1", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
            
            DB_CreateField "tblTripTrans", "AccountPaym", adVarWChar, adColNullable, 55, , "      ", False, True, , True
            DB_CreateField "tblTripTrans", "NoteSerial", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True
        End If
        

            
        If DB_CreateTable("tblTripTrans2", True, "ID", True) = True Then
            DB_CreateField "tblTripTrans2", "TravID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "TripNo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "TripDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "tblTripTrans2", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "CusID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "Typed", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "Value", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "FromID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "ToID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "CarTypeID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "CarID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "EmpID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "Remarks", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans2", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "QtyDownload", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblTripTrans2", "QtyDischarge", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblTripTrans2", "CarType1", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "CardNO", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "tblTripTrans2", "CardNO2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "tblTripTrans2", "TypeTrans", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "ShipID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "LeaderName", adVarWChar, adColNullable, 255, , "      ", False, True, , True
            
            DB_CreateField "tblTripTrans2", "notesallid", adDouble, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans2", "RecNo", adInteger, adColNullable, , , , False, True
            DB_CreateField "tblTripTrans2", "Weight", adCurrency, adColNullable, , , , False, True
            DB_CreateField "tblTripTrans2", "Price", adCurrency, adColNullable, , , , False, True
            DB_CreateField "tblTripTrans2", "TotalValue", adDouble, adColNullable, , , "    ", False, True
        End If



        
        If DB_CreateTable("tblTripTrans3", True, "ID", True) = True Then
            DB_CreateField "tblTripTrans3", "TravID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans3", "TripNo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans3", "TripDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "tblTripTrans3", "TripNo", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans3", "CarID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "tblTripTrans3", "BoardNO", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans3", "Remarks", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
            DB_CreateField "tblTripTrans3", "Price", adCurrency, adColNullable, , , , False, True
            
        End If



DB_CreateField "TblReCostCalc", "ItemID", adInteger, adColNullable, , , "  ", False, True



 DB_CreateField "Notes", "OfficeValueNet", adDouble, adColNullable, , , "  ", False, True
 DB_CreateField "Notes", "OfficeValueDiscAdd", adDouble, adColNullable, , , "  ", False, True
 DB_CreateField "Notes", "AddValue", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TblOptions", "returnnotcreatvoucher", adBoolean, adColNullable, , , "", False, True




DB_CreateField "TblPrintBarCode", "code128", adVarWChar, adColNullable, 20, , "C?C??   ", False, True, , True



DB_CreateField "TblOptions", "WaiverSetByContract", adBoolean, adColNullable, , , "", False, True

      DB_CreateField "TblItemsUnits", "ForUnit", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemsUnits", "MethodCalc", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItemsUnits", "PartItemQty", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "Transaction_Details", "OUTR", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "INR", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "TblOptions", "IsGeometricProportions", adBoolean, adColNullable, , , "", False, True




 If DB_CreateTable("TblOffline", True, "ID", True) = True Then
    DB_CreateField "TblOffline", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblOffline", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
    DB_CreateField "TblOffline", "PosName", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
    DB_CreateField "TblOffline", "CountItems", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountSales", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountSalesReturn", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountPurchase", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountPurchaseReturn", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountRec", adInteger, adColNullable, , , ""
 End If

DB_CreateField "Groups", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblUnites", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblItemLoc", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblItems", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True

DB_CreateField "TblItemProductLine", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblItemsAttach", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "ItemsPrice", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "ItemsParts", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblItemsUnits", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "Notes", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblTransactionPayments", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TransactionValueAdded", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "Transaction_Details", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "Transactions", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblMultuPayment", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True



DB_CreateField "TblOptions", "IsSomeItemWeight", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "TblOptions", "FromNo", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "OrNo", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "CodeFrom", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "CodeTo", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "WeightFrom", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "WeightTo", adCurrency, adColNullable, , , "  ", False, True
DB_CreateField "TblOptions", "IsMergeVat", adBoolean, adColNullable, , , " ", False, True

add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 9098 ,'“Ì«œ… Ê‰ﬁ’ ›Ï ‰ﬁœÌ… «·Œ“‰… ' ,'      Increase and a decrease in cash' ", "NotesType", 9098

DB_CreateField "TblOptions", "IsSerialByUserTrans", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "IsSerialByUserVouch", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblOptions", "NoOFDigitUserTrans", adInteger, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "NoOFDigitUserVouc", adInteger, adColNullable, , , "", False, True

DB_CreateField "Transactions", "OldNoteSerial", adVarWChar, adColNullable, , , "", False, True
DB_CreateField "Transactions", "OldNoteSerial1", adVarWChar, adColNullable, , , "", False, True
DB_CreateField "Transactions", "OldNoteId", adInteger, adColNullable, , , "", False, True
DB_CreateField "Transactions", "OldTransaction_ID", adInteger, adColNullable, , , "", False, True


DB_CreateField "TblOffline", "EndTime", adVarWChar, adColNullable, , , "", False, True
DB_CreateField "TblOffline", "StartTime", adVarWChar, adColNullable, , , "", False, True



DB_CreateField "Transaction_Details", "UnitPrice", adDouble, adColNullable, , , "    ", False, True

  If DB_CreateTable("TblHandWages", True, "ID", False) = True Then
                      DB_CreateField "TblHandWages", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
                        DB_CreateField "TblHandWages", "BranchID", adInteger, adColNullable, , , ""
                        
                       DB_CreateField "TblHandWages", "OrDer_no", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "OrDer_no2", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "CBoBasedON", adInteger, adColNullable, 255, , "C?C??   ", False, True, , True
                       
                       DB_CreateField "TblHandWages", "Total2", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "Total", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "VatYou", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "Vat2", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "Net", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "DiscValue", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "DiscPercent", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages", "UserID", adInteger, adColNullable, , , ""
                      DB_CreateField "TblHandWages", "Remarks", adVarWChar, adColNullable, 4000, , "?EE C?ECI??    ", False
                       
                   
                        
                       
    End If
    
    
          If DB_CreateTable("TblHandWages2", True, "ID", True) = True Then
                        DB_CreateField "TblHandWages2", "MasterID", adInteger, adColNullable, , , ""
                       DB_CreateField "TblHandWages2", "Price", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
                       DB_CreateField "TblHandWages2", "Name", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
                       
                   
                        
                       
    End If


    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 1100 ,'«·«ÃÊ— «·ÌœÊÌ…    ' ,'      Moning Items Between Inv' ", "NotesType", 1100
    
        DB_CreateField "TblHandWages", "NoteID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblHandWages", "NoteSerial", adVarWChar, adColNullable, 255, , "?EE C?ECI??    ", False
    DB_CreateField "TblHandWages", "NoteSerial1", adVarWChar, adColNullable, 255, , "?EE C?ECI??    ", False
 DB_CreateField "TblOptions", "AllowRepeateCar", adBoolean, adColNullable, , , "", False, True
 
 DB_CreateField "TblOptions", "ProvisionsByıEQuipments", adBoolean, adColNullable, , , "", False, True
 
  



DB_CreateField "TblVATAvowal", "TxtMaintCarValue1", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "TxtMaintCarValue2", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "TxtMaintCarReValue1", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "TxtMaintCarReValue2", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "TblOptions", "DontDistributeLegalACC", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "ReturnSAlesByBarcode", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblOptions", "CreatePayOrderSales", adBoolean, adColNullable, , , "", False, True





DB_CreateField "Transactions", "PayedValue2", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "Transactions", "StillValue", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True

DB_CreateField "Transactions", "NoteSerial1Cash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "Transactions", "NoteSerialCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "Transactions", "NoteIDCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True



DB_CreateField "Transactions", "DateRec", adDBTimeStamp, adColNullable, , , "      ", False, True

 

DB_CreateField "Transaction_Details", "LIPD", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "RIPD", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "LADD", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "RADD", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "LSH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "RSH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "LPRISM", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transaction_Details", "RPRISM", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

s = " Create FUNCTION GetPayValue3 (@Transaction_ID  integer,@TransType AS integer)"
s = s & "   RETURNS integer    AS    Begin"
s = s & "     RETURN ("

s = s & "   SELECT   SUM(PayedValue) AS Smatiobn"
s = s & "  From dbo.TblBillBuyPayment2"
s = s & "  Where (Transaction_ID = @Transaction_ID"
s = s & "  AND TransType = @TransType"
s = s & "  )"
s = s & "  GROUP BY Transaction_ID"
s = s & "     )"
s = s & "   End"

db_createOrUpdateFuctionSQL "GetPayValue3", s



DB_CreateField "TblTravDueK", "TotalPayed", adDouble, adColNullable, , , "      ", False, True


 '***********************************************************************************************
 
   If DB_CreateTable("TblTasks", True, "ID", False) = True Then
           DB_CreateField "TblTasks", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "TblTasks", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If
     
        DB_CreateField "TblTasks", "PercentV", adDouble, adColNullable, , , "  ", False, True
        
        
   If DB_CreateTable("TblSizesNames", True, "ID", False) = True Then
           DB_CreateField "TblSizesNames", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "TblSizesNames", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If




DB_CreateField "TblItems", "ItemRelateEmp", adBoolean, adColNullable, , , "                ", False, True



   If DB_CreateTable("TblCustomerSizes", True, "ID", True) = True Then
           DB_CreateField "TblCustomerSizes", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
           DB_CreateField "TblCustomerSizes", "SizesNamesID", adVarWChar, adColNullable, 255, , "      ", False
           DB_CreateField "TblCustomerSizes", "DateSize", adDBTimeStamp, adColNullable, , , "      ", False, True
           DB_CreateField "TblCustomerSizes", "CusId", adInteger, adColNullable, , , "  ", False, True
     End If



   If DB_CreateTable("TblJobOrders", True, "ID", False) = True Then
        DB_CreateField "TblJobOrders", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblJobOrders", "SizesNamesID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblJobOrders", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "EmpId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblJobOrders", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "CusId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "ItemID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TransactionID3", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "Noteid3", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblJobOrders", "NoteSerial13", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblJobOrders", "DateRec", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "DateRehearsal", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "RehearsalDateFinish", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "DateDelivery", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "DeliveryDateFinish", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders", "DateDeliveryAct", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        
        
        DB_CreateField "TblJobOrders", "GeneralTotal", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TotalAdd", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TotalPay", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TotalDiscPerc", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TotalDisc", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "RequiredAmount", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "PaymedValue", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders", "TotalNet", adDouble, adColNullable, , , "  ", False, True
           
     End If


  If DB_CreateTable("TblJobOrders2", True, "ID", True) = True Then
    DB_CreateField "TblJobOrders2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblJobOrders2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False

        
        DB_CreateField "TblJobOrders2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "SerID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblJobOrders2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders2", "Amount0", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "Amount2", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "Amount3", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "PercentV", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "Amount", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrders2", "DateStart", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblJobOrders2", "DateEnd", adDBTimeStamp, adColNullable, , , "      ", False, True
        
    End If



        
 
        

DB_CreateField "TblJobOrders", "NoteSerial1Cash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "TblJobOrders", "NoteSerialCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "TblJobOrders", "NoteIDCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True




   If DB_CreateTable("TblJobOrdersTasks", True, "ID", False) = True Then
        DB_CreateField "TblJobOrdersTasks", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblJobOrdersTasks", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
           
     End If
        DB_CreateField "TblJobOrdersTasks", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblJobOrdersTasks", "UserID", adInteger, adColNullable, , , "  ", False, True


If DB_CreateTable("TblJobOrdersTasks2", True, "ID", True) = True Then
    DB_CreateField "TblJobOrdersTasks2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblJobOrdersTasks2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False

        
        DB_CreateField "TblJobOrdersTasks2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "JobOrdersNo", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "EmpID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "JobOrdersNo", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "Hours", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "Total", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "PercentV", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblJobOrdersTasks2", "Amount", adDouble, adColNullable, , , "  ", False, True
        
        
    End If


'***********************************************************************************************


 DB_CreateField "TblJobOrders", "TotalAfterVat", adDouble, adColNullable, , , "  ", False, True
DB_CreateField "TblJobOrders", "Vat", adDouble, adColNullable, , , "  ", False, True
DB_CreateField "TblJobOrders", "VatYou", adDouble, adColNullable, , , "  ", False, True
DB_CreateField "TblJobOrders2", "Amount", adDouble, adColNullable, , , "  ", False, True


DB_CreateField "TblStudCalling", "Arboun", adDouble, adColNullable, , , "  ", False, True




   If DB_CreateTable("tblReservationType", True, "ID", False) = True Then
           DB_CreateField "tblReservationType", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblReservationType", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If




If DB_CreateTable("TblStudCalling2", True, "ID", True) = True Then
        
        DB_CreateField "TblStudCalling2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblStudCalling2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False

        
        DB_CreateField "TblStudCalling2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblStudCalling2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblStudCalling2", "EmpID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblStudCalling2", "ReservationTypeCode", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblStudCalling2", "Hours", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblStudCalling2", "PeriodT", adDouble, adColNullable, , , "  ", False, True
        
        
    End If








        
   If DB_CreateTable("TblAppointmentlist", True, "ID", False) = True Then
        DB_CreateField "TblAppointmentlist", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblAppointmentlist", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "TblAppointmentlist", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblAppointmentlist", "UserID", adInteger, adColNullable, , , "  ", False, True
    
     End If



If DB_CreateTable("TblAppointmentlist2", True, "ID", True) = True Then
        
        DB_CreateField "TblAppointmentlist2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblAppointmentlist2", "Timer", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblAppointmentlist2", "ReservNo", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "ServiceNo", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblAppointmentlist2", "minutes", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblAppointmentlist2", "Hours", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblAppointmentlist2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblAppointmentlist2", "CusID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "EmpID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "ReservationTypeCode", adInteger, adColNullable, , , "  ", False, True
        
        'DB_CreateField "TblAppointmentlist2", "Hours", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblAppointmentlist2", "PeriodT", adDouble, adColNullable, , , "  ", False, True
        
        
    End If



DB_CreateField "TblStudCalling2", "ItemID", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblStudCalling2", "HoursT", adVarWChar, adColNullable, 255, , "      ", False

DB_CreateField "TblStudCalling", "Arboun", adDouble, adColNullable, , , "  ", False, True




   If DB_CreateTable("TblEmpItemsTrans", True, "ID", False) = True Then
        DB_CreateField "TblEmpItemsTrans", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblEmpItemsTrans", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "TblEmpItemsTrans", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblEmpItemsTrans", "UserID", adInteger, adColNullable, , , "  ", False, True
    
     End If









If DB_CreateTable("TblEmpItemsTrans2", True, "ID", True) = True Then
        
        DB_CreateField "TblEmpItemsTrans2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpItemsTrans2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpItemsTrans2", "ItemID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpItemsTrans2", "EmpID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpItemsTrans2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        
        
        
    End If






DB_CreateField "TblOptions", "TripnotUploadExpenses", adBoolean, adColNullable, , , "", False, True
DB_CreateField "tblItems", "PeriodService", adDouble, adColNullable, , , "    ", False, True
   
   
DB_CreateField "TblOptions", "IsBarCodeByUnit", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblItemsUnits", "barCodeNo2", adVarWChar, adColNullable, 255, , " ", False, True, , True


DB_CreateField "TblAppointmentlist2", "ItemID", adInteger, adColNullable, , , "    ", False, True




DB_CreateField "TblContract", "WaterElecValueInVAT", adInteger, adColNullable, , , " ???    ", False, True

 


DB_CreateField "TblItems", "ServiceColor", adBigInt, adColNullable, , , " ???    ", False, True

 

DB_CreateField "TblJobOrders", "TransactionID1", adBigInt, adColNullable, , , " ???    ", False, True


DB_CreateField "TblJobOrders", "NoteSerial11", adVarWChar, adColNullable, 50, , "  ", False, True, , True


DB_CreateField "TblOptions", "ExpensesByQtyOnly", adBoolean, adColNullable, , , "", False, True






   DB_CreateField "TBLGeneralFundReceipt", "BoxManID", adInteger, adColNullable, , , "  ", False, True
DB_updateField "TBLGeneralFRJoin", "NoteSerial", "BigINT"



DB_CreateField "GroupsCustomers", "Account_Code", adVarWChar, adColNullable, 255, , "", False, True

DB_CreateField "TBLGeneralFundReceipt", "NoteID", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TBLGeneralFundReceipt", "NoteID", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TBLGeneralFundReceipt", "NoteSerial", adBigInt, adColNullable, , , "  ", False, True
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 1063 ,'”‰œ «·ﬁ»÷ «·⁄«„   ' ,'      Moning Items Between Inv' ", "NotesType", 1063

Dim SPH As Double
Dim CLY As Double
Dim aix As Double

If DB_CreateTable("SPHTable", True, "ID", True) = True Then
        
 DB_CreateField "SPHTable", "SPH", adDouble, adColNullable, , , "    ", False, True
 
 
         End If
'*******************************************
If DB_CreateTable("CLYTable", True, "ID", True) = True Then
        
 DB_CreateField "CLYTable", "CLY", adDouble, adColNullable, , , "    ", False, True
        
         
         End If
         DB_CreateField "SPHTable", "SPHT", adVarWChar, adColNullable, 255, , "", False, True
         DB_CreateField "CLYTable", "CLYT", adVarWChar, adColNullable, 255, , "", False, True
'*******************************************
If DB_CreateTable("aixTable", True, "ID", True) = True Then
        
 DB_CreateField "aixTable", "aix", adDouble, adColNullable, , , "    ", False, True
         End If
'*******************************************

For SPH = 20 To -20 Step -0.25
add_record_to_table "SPHTable", "SPH", Format(CStr(SPH), "0.00"), "sph", CDbl(SPH)
 

Next SPH



For CLY = 8 To -8 Step -0.25
add_record_to_table "CLYTable", "CLY", CStr(CLY), "CLY", CDbl(CLY)
Next CLY

 


Cn.Execute "update SPHTable set   spht= concat('+' ,  FORMAT(sph, 'N', 'en-us') )  where sph>0 and spht is null "
Cn.Execute "update SPHTable set spht= FORMAT(sph, 'N', 'en-us') where sph<0 and spht is null "

Cn.Execute "update CLYTable set   CLYt= concat('+' ,  FORMAT(CLY, 'N', 'en-us') )  where CLY>0 and CLYt is null "
Cn.Execute "update CLYTable set CLYt=  FORMAT(CLY, 'N', 'en-us') where CLY <0 and CLYt is null "




For aix = 0 To 180
add_record_to_table "aixTable", "aix", CStr(aix), "aix", CDbl(aix)
Next aix




   If DB_CreateTable("tblPaymentClass", True, "ID", False) = True Then
           DB_CreateField "tblPaymentClass", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblPaymentClass", "NameE", adVarWChar, adColNullable, 255, , "      ", False
     End If


 If DB_CreateTable("TblTripReg", True, "ID", False) = True Then
        DB_CreateField "TblTripReg", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTripReg", "LocationsName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTripReg", "CarName", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTripReg", "CustName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTripReg", "PhoneCust", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTripReg", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTripReg", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTripReg", "PayMentType", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTripReg", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTripReg", "PaymentClassID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTripReg", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTripReg", "Value", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTripReg", "VAt22", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTripReg", "TotalWithVat2", adDouble, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTripReg", "DateRec", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        
    End If



   DB_CreateField "TBLGeneralFundReceipt", "BoxManID", adInteger, adColNullable, , , "  ", False, True
DB_updateField "TBLGeneralFRJoin", "NoteSerial", "BigINT"



DB_CreateField "GroupsCustomers", "Account_Code", adVarWChar, adColNullable, 255, , "", False, True

DB_CreateField "TBLGeneralFundReceipt", "NoteID", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TBLGeneralFundReceipt", "NoteID", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TBLGeneralFundReceipt", "NoteSerial", adBigInt, adColNullable, , , "  ", False, True
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 1063 ,'”‰œ «·ﬁ»÷ «·⁄«„   ' ,'      Moning Items Between Inv' ", "NotesType", 1063

DB_CreateField "TblTripReg", "NoteSerial1", adVarWChar, adColNullable, 255, , "?EE C?ECI??    ", False



DB_CreateField "TblTripReg", "CusId", adInteger, adColNullable, , , " ???    ", False, True


DB_CreateField "TblTripReg", "AmountLater", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblTripReg", "AmountCash", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblTripReg", "PayType", adInteger, adColNullable, , , " ???    ", False, True



DB_CreateField "TblTripReg", "CusId", adInteger, adColNullable, , , " ???    ", False, True


DB_CreateField "TblTripReg", "AmountLater", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblTripReg", "AmountCash", adInteger, adColNullable, , , " ???    ", False, True
DB_CreateField "TblTripReg", "AmountVisa", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "TblTripReg", "PayType", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "tblPaymentClass", "IsBoardNo", adBoolean, adColNullable, , , " ???    ", False, True


DB_CreateField "tblPaymentClass", "BoardNo", adBoolean, adColNullable, , , " ???    ", False, True

DB_CreateField "tblPaymentClass", "BoardNo", adBoolean, adColNullable, , , " ???    ", False, True


    DB_CreateField "TblTripReg", "nBoardNo", adVarWChar, adColNullable, 255, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "BoardNo", adVarWChar, adColNullable, 255, , "  ", False, True, , True
    
    DB_CreateField "TblTripReg", "txtLetter1", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "txtLetter2", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "txtLetter3", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "txtLetter4", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    
    
    
    DB_CreateField "TblTripReg", "ntxtLetter1", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "ntxtLetter2", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "ntxtLetter3", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    DB_CreateField "TblTripReg", "ntxtLetter4", adVarWChar, adColNullable, 10, , "  ", False, True, , True
    

DB_CreateField "TblTripReg", "txtNum1", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "txtNum2", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "txtNum3", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "txtNum4", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "ntxtNum1", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "ntxtNum2", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "ntxtNum3", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblTripReg", "ntxtNum4", adInteger, adColNullable, , , "  ", False, True




DB_CreateField "TblCustemers", "IsNotCommission", adBoolean, adColNullable, , , " ???    ", False, True

DB_CreateField "tblPaymentClass", "ServiceColor", adBigInt, adColNullable, , , " ???    ", False, True


       
   If DB_CreateTable("TblEmpData", True, "ID", False) = True Then
        DB_CreateField "TblEmpData", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblEmpData", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblEmpData", "startDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
            DB_CreateField "TblEmpData", "BranchId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpData", "HafizaNo", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "EmpName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "LocationsName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "salary", adDouble, adColNullable, , , "  ", False, True
       DB_CreateField "TblEmpData", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
       DB_CreateField "TblEmpData", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblEmpData", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpData", "IsEmp", adInteger, adColNullable, , , "  ", False, True
    
     End If


        

    If DB_CreateTable("TblEmpDataFingerPrint", True, "ID", True) = True Then
        
        DB_CreateField "TblEmpDataFingerPrint", "EmpId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpDataFingerPrint", "FingerPrint", adVarWChar, adColNullable, 4000, , "      ", False
    
     End If
       


        
   If DB_CreateTable("TblEmpData", True, "ID", False) = True Then
        DB_CreateField "TblEmpData", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblEmpData", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblEmpData", "startDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
            DB_CreateField "TblEmpData", "BranchId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpData", "HafizaNo", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "EmpName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "LocationsName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpData", "salary", adDouble, adColNullable, , , "  ", False, True
       DB_CreateField "TblEmpData", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
       DB_CreateField "TblEmpData", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblEmpData", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpData", "IsEmp", adInteger, adColNullable, , , "  ", False, True
    
     End If


DB_CreateField "TblEmpData", "FingerStatus", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "TblEmpData", "FingerPrint", adVarWChar, adColNullable, 4000, , "      ", False
 


    DB_CreateField "TblEmpData", "FinferPrint", adVarWChar, adColNullable, 4000, , "      ", False


DB_CreateField "TblEmpData", "Photo2", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True

DB_CreateField "TblEmpDataInOut", "Hours", adVarWChar, adColNullable, 255, , "      ", False
        
        


   If DB_CreateTable("TblEmpDataInOut", True, "ID", False) = True Then
        DB_CreateField "TblEmpDataInOut", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblEmpDataInOut", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblEmpDataInOut", "startDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblEmpDataInOut", "BranchId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpDataInOut", "HafizaNo", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpDataInOut", "EmpId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblEmpDataInOut", "EmpName", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblEmpDataInOut", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblEmpDataInOut", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblEmpDataInOut", "UserID", adInteger, adColNullable, , , "  ", False, True
        
    
     End If


DB_CreateField "TblOptions", "ShowPrinterDialoge", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblEmpData", "MobileNO", adVarWChar, adColNullable, 255, , "      ", False

   If DB_CreateTable("TblIqarDiscountTrans", True, "ID", False) = True Then
        DB_CreateField "TblIqarDiscountTrans", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblIqarDiscountTrans", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "TblIqarDiscountTrans", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblIqarDiscountTrans", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblIqarDiscountTrans", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        
     End If


DB_CreateField "TblContract", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True

If DB_CreateTable("TblIqarDiscountTrans2", True, "ID", True) = True Then
        
        DB_CreateField "TblIqarDiscountTrans2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblIqarDiscountTrans2", "BranchID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "Iqar", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "unittype", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "UnitNo", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblIqarDiscountTrans2", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    End If

If DB_CreateTable("tblRestsTypes", True, "ID", False) = True Then
           DB_CreateField "tblRestsTypes", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblRestsTypes", "NameE", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblRestsTypes", "ServiceColor", adBigInt, adColNullable, , , " ???    ", False, True
     End If
   
   
   
   
   
   
    If DB_CreateTable("tblRestsSiftTrans", True, "ID", False) = True Then
        DB_CreateField "tblRestsSiftTrans", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "tblRestsSiftTrans", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "tblRestsSiftTrans", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "tblRestsSiftTrans", "UserID", adInteger, adColNullable, , , "  ", False, True
        
        
        
        
     End If
     
     
     DB_CreateField "TblStudCalling2", "Status", adVarWChar, adColNullable, 255, , "      ", False
     If DB_CreateTable("tblRestsSiftTrans2", True, "ID", True) = True Then
        
        DB_CreateField "tblRestsSiftTrans2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "tblRestsSiftTrans2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "tblRestsSiftTrans2", "RestsTypesID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "tblRestsSiftTrans2", "ShiftTypeID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "tblRestsSiftTrans2", "EmpID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "tblRestsSiftTrans2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "tblRestsSiftTrans2", "FromTime", adVarWChar, adColNullable, 20, , "      ", False, True, , True
        DB_CreateField "tblRestsSiftTrans2", "ToTime", adVarWChar, adColNullable, 20, , "      ", False, True, , True
 
        
        
        DB_CreateField "tblRestsSiftTrans2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "tblRestsSiftTrans2", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    End If

   
   If DB_CreateTable("TblIqarDiscountTrans", True, "ID", False) = True Then
        DB_CreateField "TblIqarDiscountTrans", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblIqarDiscountTrans", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
               DB_CreateField "TblIqarDiscountTrans", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblIqarDiscountTrans", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblIqarDiscountTrans", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        
     End If


DB_CreateField "TblContract", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True

If DB_CreateTable("TblIqarDiscountTrans2", True, "ID", True) = True Then
        
        DB_CreateField "TblIqarDiscountTrans2", "SerID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblIqarDiscountTrans2", "BranchID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "Iqar", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "unittype", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "UnitNo", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblIqarDiscountTrans2", "DiscountPercent", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblIqarDiscountTrans2", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    End If

   DB_CreateField "TblEmpDataFingerPrint", "FingerPrint2", adVarWChar, adColNullable, 4000, , "      ", False
        DB_CreateField "TblEmpDataFingerPrint", "FingerPrint3", adVarWChar, adColNullable, 4000, , "      ", False
        DB_CreateField "TblEmpDataFingerPrint", "FingerPrint4", adVarWChar, adColNullable, 4000, , "      ", False
        DB_CreateField "TblEmpDataFingerPrint", "FingerPrint5", adVarWChar, adColNullable, 4000, , "      ", False

DB_CreateField "tblPaymentClass", "IsBoardNoHide", adBoolean, adColNullable, , , " ???    ", False, True

DB_CreateField "TblContract", "DiscountvaLUE", adDouble, adColNullable, , , "  ", False, True
DB_CreateField "TblCustemers", "NoComm", adInteger, adColNullable, , , " ???    ", False, True

                  DB_CreateField "notes_all", "CurrncyID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "notes_all", "rate", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblStudCalling", "NoteSerialCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "TblStudCalling", "NoteIDCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True



DB_CreateField "TblStudCalling", "NoteSerialCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "TblStudCalling", "NoteIDCash", adBigInt, adColNullable, 255, , "C?C??   ", False, True, , True

DB_CreateField "notes_all", "Vat2", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "notes_all", "Net", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
  DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "VATYou", adDouble, adColNullable, 255, , "C?C??   ", False, True, , True
  
    DB_CreateField "TblVATAvowal", "FaBuy", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblVATAvowal", "FaBuy2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblVATAvowal", "FaBuy3", adDouble, adColNullable, , , "    ", False, True
  

 
DB_CreateField "TblOptions", "NOOFPRINTCOPIESSALES", adInteger, adColNullable, , , "  Â· Ì „  ÕœÌœ ⁄œœ «·Œ«‰«  ›Ì „” Œ·’«  «·„‘«—Ì⁄  ", False, True
DB_CreateField "TblOptions", "AllowUnbalncedByBranchAccount", adBoolean, adColNullable, , , " ”œ«œ „ ⁄œœ ", False, True

DB_CreateField "TblTripReg", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblEmpDataInOut", "Copied", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "TblEmpDataFingerPrint", "HafizaNo", adVarWChar, adColNullable, 255, , "      ", False
DB_CreateField "TblTripReg", "OldNoteSerial1", adVarWChar, adColNullable, 255, , "?EE C?ECI??    ", False

DB_CreateField "TblTripReg", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblTripReg", "OldId", adInteger, adColNullable, , , "    ", False, True




 If DB_CreateTable("TblOffline2", True, "ID", True) = True Then
    DB_CreateField "TblOffline2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblOffline2", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
    DB_CreateField "TblOffline2", "PosName", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
    DB_CreateField "TblOffline2", "CountItems", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline2", "CountSales", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline2", "CountSalesReturn", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline2", "CountPurchase", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline2", "CountPurchaseReturn", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline2", "CountRec", adInteger, adColNullable, , , ""
 End If



DB_CreateField "TblTripReg", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblEmpDataFingerPrint", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
DB_CreateField "TblEmpData", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True

DB_CreateField "TblEmpDataInOut", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True

DB_CreateField "TblEmpDataInOut", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblEmpDataFingerPrint", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblEmpData", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblTripReg", "Copied", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "TblEmpDataInOut", "OldId", adInteger, adColNullable, , , "    ", False, True


DB_CreateField "notes_all", "CountF", adInteger, adColNullable, , , "    ", False, True



DB_CreateField "Transactions", "IsMaxFromInvoice", adBoolean, adColNullable, , , "", False, True


DB_CreateField "TblBillBuyPayment2", "Transaction_IDReturn", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblBillBuyPayment2", "AUToJoinid", adDouble, adColNullable, , , "    ", False, True

  DB_CreateField "TblContract", "Accredit", adBoolean, adColNullable, , , "        ", False, True
  
  
 DB_CreateField "Transactions", "InsuranceCompanyid", adDouble, adColNullable, , , "    ", False, True
 DB_CreateField "Transactions", "DoctorId", adDouble, adColNullable, , , "    ", False, True
 
 
           If DB_CreateTable("tblDoctorsType", True, "ID", False) = True Then
           DB_CreateField "tblDoctorsType", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblDoctorsType", "NameE", adVarWChar, adColNullable, 255, , "      ", False
   DB_CreateField "tblDoctorsType", "PercentV", adDouble, adColNullable, , , "  ", False, True
     End If
  
  
   DB_CreateField "insurance_companies", "Descount", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "insurance_companies", "Policy", adVarWChar, adColNullable, 4000, , "      ", False
  DB_CreateField "insurance_companies", "TypeID", adDouble, adColNullable, , , "    ", False, True
  
     If DB_CreateTable("MedicalContractingType", True, "ID", False) = True Then
     DB_CreateField "MedicalContractingType", "Name", adVarWChar, adColNullable, 255, , "      ", False
     DB_CreateField "MedicalContractingType", "Namee", adVarWChar, adColNullable, 255, , "      ", False
     
     End If
     
  add_record_to_table "MedicalContractingType", "id,Namee,Name ", " 1, 'Insurance', '  √„Ì‰'    ", "id", 1
  add_record_to_table "MedicalContractingType", "id,Namee,Name ", " 2, 'Discount Card', ' ﬂ—Ê  Œ’„'    ", "id", 2
  add_record_to_table "MedicalContractingType", "id,Namee,Name ", " 3, 'Private Card', ' ﬂ—Ê  Œ«’…'    ", "id", 3
  
  
   DB_CreateField "Transactions", "ApprovalNO", adVarWChar, adColNullable, 4000, , "      ", False
  DB_CreateField "Transactions", "ApprovalValue", adDouble, adColNullable, , , "    ", False, True
  
  
  DB_CreateField "TblOptions", "SortInvoiceByEntry", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOptions", "CostProductOrderByOut", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblUsers", "CanEditOnlyPayMethod", adBoolean, adColNullable, , , "", False, True


 DB_CreateField "TblEmpIncreaseSalary", "Approved", adBoolean, adColNullable, , , "        ", False, True
  
   DB_CreateField "TblEmpIncreaseSalary", "DateIncrease", adDBTimeStamp, adColNullable, , , "      ", False, True
   DB_CreateField "TblEmpIncreaseSalary", "Remark", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
     DB_CreateField "TblEmpIncreaseSalaryDetalis", "Typeincrease", adInteger, adColNullable, , , " ???    ", False, True
   DB_CreateField "TblEmpIncreaseSalaryDetalis", "TypeValue", adDouble, adColNullable, , , "    ", False, True
     DB_CreateField "TblEmpIncreaseSalaryDetalis", "CurrValue", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblEmpIncreaseSalaryDetalis", "IncreaseValue", adDouble, adColNullable, , , "    ", False, True
              
              

   If DB_CreateTable("TblTipDates", True, "ID", False) = True Then
        DB_CreateField "TblTipDates", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDates", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDates", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates", "UserID", adInteger, adColNullable, , , "  ", False, True
        
        
        
        DB_CreateField "TblTipDates", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        
        DB_CreateField "TblTipDates", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
           
     End If


   If DB_CreateTable("TblTipDatesReg", True, "ID", False) = True Then
        DB_CreateField "TblTipDatesReg", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDatesReg", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDatesReg", "TipDatesID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "optWayTow", adInteger, adColNullable, , , "  ", False, True
        
        
        
        DB_CreateField "TblTipDatesReg", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        
        DB_CreateField "TblTipDatesReg", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "TotalTripReg", adDouble, adColNullable, , , "  ", False, True
        
          
           
     End If


  If DB_CreateTable("TblTipDates2", True, "ID", True) = True Then
    DB_CreateField "TblTipDates2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDates2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False

        
        DB_CreateField "TblTipDates2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SerID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDates2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SetCount", adDouble, adColNullable, , , "  ", False, True
        
        
        DB_CreateField "TblTipDates2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        
        DB_CreateField "TblTipDates2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True

        
    End If




  If DB_CreateTable("TblTipDatesReg2", True, "ID", True) = True Then
    DB_CreateField "TblTipDatesReg2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDatesReg2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False

        
        DB_CreateField "TblTipDatesReg2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "MasterID2", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SerID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDatesReg2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SetCount", adDouble, adColNullable, , , "  ", False, True
        
        
        DB_CreateField "TblTipDatesReg2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        
        DB_CreateField "TblTipDatesReg2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True

        
    End If




Dim MySQL As String
       MySQL = " SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
 MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
 MySQL = MySQL & "                      dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType,"
 MySQL = MySQL & "                      dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Parent_Account_Code,"
 MySQL = MySQL & "                      dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center,"
 MySQL = MySQL & "                      dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1,"
 MySQL = MySQL & "                      dbo.TblNotesTypes.NotesTypeNamee, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id,"
 MySQL = MySQL & "                      dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid,"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo "
 MySQL = MySQL & "    FROM         dbo.TblAqar RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBranchesData INNER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers INNER JOIN"
 MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON"
 MySQL = MySQL & "                      dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.Notes LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
 MySQL = MySQL & "     Where (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)"
  db_createOrUpdateviewSQL "RptLedger_Sub", MySQL
  
  
UpdateDataBasePart26
   

 
End Function
Function UpdateDataBasePart26()

    On Error Resume Next
    Dim New_View As String
    Dim s        As String
    DB_CreateField "TblAqrOwin", "[Select]", adBoolean, adColNullable, , , "                ", False, True
    If DB_CreateTable("TblTipDates", True, "ID", False) = True Then
        DB_CreateField "TblTipDates", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDates", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDates", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates", "UserID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        DB_CreateField "TblTipDates", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
           
    End If

    If DB_CreateTable("TblTipDatesReg", True, "ID", False) = True Then
        DB_CreateField "TblTipDatesReg", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDatesReg", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDatesReg", "TipDatesID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "optWayTow", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        DB_CreateField "TblTipDatesReg", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "TotalTripReg", adDouble, adColNullable, , , "  ", False, True
           
    End If

    DB_CreateField "TblTipDatesReg", "CountryID2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "CountryID12", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "FromDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblTipDatesReg", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblTipDatesReg", "CountPerson", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "CountPerson2", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblTipDatesReg", "TotalTripReg", adDouble, adColNullable, , , "  ", False, True

    DB_CreateField "TblTipDates2", "GWay", adInteger, adColNullable, , , "  ", False, True

    If DB_CreateTable("TblTipDates2", True, "ID", True) = True Then
        DB_CreateField "TblTipDates2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDates2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDates2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SerID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDates2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SetCount", adDouble, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDates2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        DB_CreateField "TblTipDates2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
        
    End If

    If DB_CreateTable("TblTipDatesReg2", True, "ID", True) = True Then
        DB_CreateField "TblTipDatesReg2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDatesReg2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblTipDatesReg2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "MasterID2", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SerID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblTipDatesReg2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SetCount", adDouble, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblTipDatesReg2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        
        DB_CreateField "TblTipDatesReg2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
        
    End If
    
    DB_CreateField "TblTipDates", "SetCount", adDouble, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDates", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg2", "GWay", adInteger, adColNullable, , , "  ", False, True
            
    If DB_CreateTable("TblTipDates", True, "ID", False) = True Then
        DB_CreateField "TblTipDates", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
            
        DB_CreateField "TblTipDates", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
            
        DB_CreateField "TblTipDates", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "BranchId", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDates", "UserID", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDates", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
        DB_CreateField "TblTipDates", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
               
    End If
    
    DB_CreateField "TblTipDatesReg", "CountPerson", adInteger, adColNullable, , , "  ", False, True
    
    If DB_CreateTable("TblTipDatesReg", True, "ID", False) = True Then
        DB_CreateField "TblTipDatesReg", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
            
        DB_CreateField "TblTipDatesReg", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            
        DB_CreateField "TblTipDatesReg", "TipDatesID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "BranchId", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDatesReg", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "optWayTow", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDatesReg", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "StartTime", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg", "StartTime2", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
        DB_CreateField "TblTipDatesReg", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg", "TotalTripReg", adDouble, adColNullable, , , "  ", False, True
               
    End If
    
    DB_CreateField "TblTipDatesReg", "CountryID2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "CountryID12", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "FromDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblTipDatesReg", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblTipDatesReg", "CountPerson", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblTipDatesReg", "CountPerson2", adInteger, adColNullable, , , "  ", False, True
    
    DB_CreateField "TblTipDatesReg", "TotalTripReg", adDouble, adColNullable, , , "  ", False, True
    
    DB_CreateField "TblTipDates2", "GWay", adInteger, adColNullable, , , "  ", False, True
    
    If DB_CreateTable("TblTipDates2", True, "ID", True) = True Then
        DB_CreateField "TblTipDates2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDates2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
            
        DB_CreateField "TblTipDates2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SerID", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDates", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDates2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
            
        DB_CreateField "TblTipDates2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "BranchId", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDates2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "SetCount", adDouble, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDates2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDates2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
        DB_CreateField "TblTipDates2", "SetCount", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDates2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
            
    End If
    
    If DB_CreateTable("TblTipDatesReg2", True, "ID", True) = True Then
        DB_CreateField "TblTipDatesReg2", "SerID", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblTipDatesReg2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
            
        DB_CreateField "TblTipDatesReg2", "MasterID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "MasterID2", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TasksID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SerID", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDatesReg2", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "ToDate2", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTipDatesReg2", "DateTrip", adDBTimeStamp, adColNullable, , , "      ", False, True
            
        DB_CreateField "TblTipDatesReg2", "CarId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "BranchId", adInteger, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDatesReg2", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "PeriodTrip", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "SetCount", adDouble, adColNullable, , , "  ", False, True
            
        DB_CreateField "TblTipDatesReg2", "CountryID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "CountryID1", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblTipDatesReg2", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "TimeOut", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
        DB_CreateField "TblTipDatesReg2", "PeriodTripM", adVarWChar, adColNullable, 50, , "      ", False, True, , True
            
        DB_CreateField "TblTipDatesReg2", "TravelPrice", adDouble, adColNullable, , , "  ", False, True
            
    End If
    
    DB_CreateField "TblOptions", "TransferNotInvItemDef", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblUsers", "CanTransferItemDef", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblOptions", "CustMobNoMandatory", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblRestsSiftTrans", "EmpID", adInteger, adColNullable, , , ""
    DB_CreateField "tblRestsSiftTrans", "ShiftID", adInteger, adColNullable, , , ""
    DB_CreateField "tblRestsSiftTrans", "RestsTypesID", adInteger, adColNullable, , , ""

    DB_CreateField "tblRestsSiftTrans", "FromTime", adVarWChar, adColNullable, 20, , "", False, True, , True
    DB_CreateField "tblRestsSiftTrans", "ToTime", adVarWChar, adColNullable, 20, , "", False, True, , True

    DB_CreateField "tblRestsSiftTrans", "FromDate", adDBTimeStamp, adColNullable, 8, , "     ", False
    DB_CreateField "tblRestsSiftTrans", "ToDate", adDBTimeStamp, adColNullable, 8, , "     ", False
    
    DB_CreateField "TblDefComItemData", "ItemID5", adInteger, adColNullable, , , ""
    DB_CreateField "TblDefComItemData", "CountItem2", adInteger, adColNullable, , , ""
    DB_CreateField "TblDefComItemData", "CountItem5", adInteger, adColNullable, , , ""

    DB_CreateField "TblUsers", "CanPrintMultiSales", adBoolean, adColNullable, , , "", False, True
 
    If DB_CreateTable("TblRowsEstimated", True, "ID", False) = True Then
        DB_CreateField "TblRowsEstimated", "ID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "AuthoOrder", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "AuthoOrderID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "CarModelID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblRowsEstimated", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "UserID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "CarTypeID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "PlateNo", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "ColorID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "CarID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "CusID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "YearFact", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "ClientCode", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblRowsEstimated", "ClientName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblRowsEstimated", "Shaseh", adVarWChar, adColNullable, 400, , "      ", False, True, , True
                 
        DB_CreateField "TblRowsEstimated", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
    End If
          
    DB_CreateField "TblRowsEstimated", "Discount", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "TotalAfterDisc", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "Vat", adDouble, adColNullable, , , " ???    ", False, True

    DB_CreateField "TblRowsEstimated2", "VatValue", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "Net", adDouble, adColNullable, , , " ???    ", False, True
 
    DB_CreateField "TblRowsEstimated", "VatValue", adDouble, adColNullable, , , " ???    ", False, True
          
    DB_CreateField "TblRowsEstimated2", "GroupID", adInteger, adColNullable, , , " ???    ", False, True
             
    DB_CreateField "TblRowsEstimated2", "Discount", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "TotalAfterDisc", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "Vat", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "Net", adDouble, adColNullable, , , " ???    ", False, True
              
    If DB_CreateTable("TblRowsEstimated2", True, "ID", False) = True Then
        DB_CreateField "TblRowsEstimated2", "ID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "GroupID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "ItemID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "UnitID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "ShowQty", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "UnitPrice", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "Total", adDouble, adColNullable, , , " ???    ", False, True
                   
        DB_CreateField "TblRowsEstimated2", "Discount", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "TotalAfterDisc", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "Vat", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "Net", adDouble, adColNullable, , , " ???    ", False, True
                   
        DB_CreateField "TblRowsEstimated2", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
                   
        DB_CreateField "TblRowsEstimated2", "ItemName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblRowsEstimated2", "FullCode", adVarWChar, adColNullable, 400, , "      ", False, True, , True
                   
        DB_CreateField "TblRowsEstimated2", "Rem", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If
                          
    DB_CreateField "TblRowsEstimated", "Discount", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "TotalAfterDisc", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "Vat", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "Net", adDouble, adColNullable, , , " ???    ", False, True
          
    DB_CreateField "TblRowsEstimated2", "GroupID", adInteger, adColNullable, , , " ???    ", False, True
             
    DB_CreateField "TblRowsEstimated2", "Discount", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "TotalAfterDisc", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "Vat", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "VatValue", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "Net", adDouble, adColNullable, , , " ???    ", False, True
 
    DB_CreateField "TblRowsEstimated", "VatValue", adDouble, adColNullable, , , " ???    ", False, True
                    
    DB_CreateField "TblRowsEstimated", "AuthoOrder", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "AuthoOrderID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated2", "GroupID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "CarTypeID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "PlateNo", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "CarModelID", adInteger, adColNullable, , , " ???    ", False, True
                 
    DB_CreateField "TblRowsEstimated", "ColorID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "CarID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "CusID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "YearFact", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "ClientCode", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblRowsEstimated", "ClientName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblRowsEstimated", "Shaseh", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblRowsEstimated2", "GroupID", adInteger, adColNullable, , , " ???    ", False, True
    If DB_CreateTable("TblRowsEstimated", True, "ID", False) = True Then
        DB_CreateField "TblRowsEstimated", "ID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblRowsEstimated", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated", "UserID", adInteger, adColNullable, , , " ???    ", False, True
    End If
              
    If DB_CreateTable("TblRowsEstimated2", True, "ID", False) = True Then
        DB_CreateField "TblRowsEstimated2", "ID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "ItemID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "UnitID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "ShowQty", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "UnitPrice", adDouble, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblRowsEstimated2", "Total", adDouble, adColNullable, , , " ???    ", False, True
                   
        DB_CreateField "TblRowsEstimated2", "ItemName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblRowsEstimated2", "FullCode", adVarWChar, adColNullable, 400, , "      ", False, True, , True
                   
        DB_CreateField "TblRowsEstimated2", "Rem", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "TblRowsEstimated", "DiscValue", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "DiscPercent", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "TotalAfterDiscount", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "Vatyo", adDouble, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblRowsEstimated", "Vat2", adDouble, adColNullable, , , " ???    ", False, True
            
    DB_CreateField "TblDefComItem", "order_no", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblDefComItem", "OrderID", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblDefComItem", "order_no", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblDefComItem", "OrderID", adInteger, adColNullable, , , " ???    ", False, True

    DB_CreateField "TblOptions", "CostByProduction", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblRowsEstimated", "CBoBasedON", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblRowsEstimated", "order_no", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblRowsEstimated", "orderStatus", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblRowsEstimated", "CarMeter", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblRowsEstimated", "HandWagesAmount", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblRowsEstimated", "Remarks", adVarChar, adColNullable, 4000, , "", False, True

    DB_CreateField "TblOptions", "MaintOrderCantRepeatSales", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "MaintOrderCantRepeatBillBuy", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblOptions", "PaymentMethLaterCompItem", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "ShowBalanceCustInv", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "AreaL", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "TblDefComItemData", "AreaL", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "TblHandWages", "RowsEstimatedID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "ExpensesType", "Transportation", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "TripRevenueAuto", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblOptions", "cdoSMTPServer", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "TblOptions", "cdoSendUserName", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "TblOptions", "txtFromName", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "TblOptions", "txtFromEmail", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "TblOptions", "cdoSendPassword", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "TblOptions", "cdoSMTPUseSSL", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "cdoSMTPServerPort", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblCusCsh", "Email", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "TblUsers", "CanPayWithoutPrint", adBoolean, adColNullable, , , "", False, True
                     
    DB_CreateField "TblItemShows", "Sa", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "Su", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "Mo", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "Tu", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "We", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "Th", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblItemShows", "Fr", adBoolean, adColNullable, , , "                ", False, True

    DB_CreateField "TblStore", "Account_Code0", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "TblStore", "Account_Code11", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "TblStore", "Account_Code22", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "TblStore", "Account_Code33", adVarWChar, adColNullable, 255, , " ", False, True, , True
       
    If DB_CreateTable("tblUserPermAccounts", True, "ID", True) = True Then
        DB_CreateField "tblUserPermAccounts", "UserId", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "tblUserPermAccounts", "AccountCode", adVarWChar, adColNullable, 255, , " ", False, True, , True
              
    End If
    DB_CreateField "tblUserPermAccounts", "AccountCode", adVarWChar, adColNullable, 255, , " ", False, True, , True
    DB_CreateField "Accounts", "Level", adInteger, adColNullable, , , " ???    ", False, True

    DB_CreateField "TblItems", "BrandsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "TypeItemsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "DesignID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "CollectionsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "ShapesID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "ShapesNewID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "MaterialID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblItems", "SexID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "AGEID", adInteger, adColNullable, , , "", False, True
    
    If DB_CreateTable("tblBrands", True, "ID", False) = True Then
        DB_CreateField "tblBrands", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblBrands", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblTypeItems", True, "ID", False) = True Then
        DB_CreateField "tblTypeItems", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblTypeItems", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblDesign", True, "ID", False) = True Then
        DB_CreateField "tblDesign", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblDesign", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblCollections", True, "ID", False) = True Then
        DB_CreateField "tblCollections", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblCollections", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblShapes", True, "ID", False) = True Then
        DB_CreateField "tblShapes", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblShapes", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblShapesNew", True, "ID", False) = True Then
        DB_CreateField "tblShapesNew", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblShapesNew", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblMaterial", True, "ID", False) = True Then
        DB_CreateField "tblMaterial", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblMaterial", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "sanad_numbering", "IsBreaks", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "sanad_numbering", "IsCodeByBranch", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "sanad_numbering", "IsSerialByUser", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "sanad_numbering", "Breaks", adVarWChar, adColNullable, 255, , " ", False, True, , True

    DB_CreateField "TblOptions", "IsByNewCoding", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "IsAutoNameItems", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "ACCOUNTS", "BranchID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "Transactions", "Ser", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Notes", "Ser", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblItems", "NationalityID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "ColorID1", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "isDeactivated", adInteger, adColNullable, , , " ???    ", False, True
        
    DB_CreateField "TblDefComItemDet", "IsRow", adBoolean, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemDet", "widtj", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemDet", "hight", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemDet", "Length", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemDet", "thickness", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemDet", "DO", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemDet", "DI", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemDet", "Diameter", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblEmployee", "BasicSalary", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeFood", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeMove", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeHome", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeOther", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeFixed", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeLoca", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "FeeTel", adDouble, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "TotalSalary", adDouble, adColNullable, , , "", False, True

    DB_CreateField "TblEmployee", "GroupSalary", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblEmployee", "SectorName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    
    DB_CreateField "TblItems", "ColorID1", adInteger, adColNullable, , , "", False, True

   ' Cn.Execute "ALTER TABLE TblItems  DROP COLUMN ColorID"

    sql = "    DROP FUNCTION QryItemsInventry3" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION [dbo].[QryItemsInventry3] (@fromdate datetime,@todate datetime,@StoreId AS INT=null,@ColorID AS INT=null,@ItemSize AS  NVARCHAR(255)=null ,"
    sql = sql & "  @ClassId AS INT=null , @order_no  AS  NVARCHAR(255)  =null,@CusID as float=null)"
    sql = sql & " RETURNS @XTable Table"
    sql = sql & "    ("
    sql = sql & "      Item_ID  Decimal (18,2),"
    sql = sql & "   LotNO  nvarchar(255),"
    sql = sql & " ItemCode     nvarchar(255)  ,"
    sql = sql & "     ItemName  nvarchar(4000)     ,"
    sql = sql & "   openingValue  Decimal (18,2),"
    sql = sql & "    inputvalue  Decimal (18,2),"
    sql = sql & "   outputValue Decimal(18, 2)"
    sql = sql & " )"
    sql = sql & "  AS"
    sql = sql & " Begin"
    sql = sql & " INSERT  @XTable"
    sql = sql & " Select      Item_ID,  LotNO, ItemCode, ItemName,  Sum(DEV_Value1) as openingValue,Sum(DEV_Value2) as inputvalue , Sum(DEV_Value3) as outputValue"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & " SELECT"
    sql = sql & " Item_ID,  null as LotNO, ItemCode, ItemName,"
    sql = sql & " DEV_Value1=Case"
    sql = sql & "  When  (dbo.TransactionTypes.StockEffect = 1)  and (dbo.Transactions.Transaction_Type=3)   Then  (Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & "  END,"
    sql = sql & "  DEV_Value2=Case"
    sql = sql & "  When  (dbo.TransactionTypes.StockEffect = 1)  and (dbo.Transactions.Transaction_Type<>3)   Then  (Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & "  End"
    sql = sql & "   ,"
    sql = sql & " DEV_Value3=Case"
    sql = sql & " When    (dbo.TransactionTypes.StockEffect = -1)   Then ( Quantity * dbo.TransactionTypes.StockEffect)"
    sql = sql & " Else 0"
    sql = sql & " End"
    sql = sql & "  FROM         dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    sql = sql & " INNER JOIN  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    sql = sql & " INNER JOIN    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    sql = sql & " where dbo.Transactions.Transaction_Date>=@fromdate"
    sql = sql & " and dbo.Transactions.Transaction_Date <=@todate"
    sql = sql & " and  Storeid=isnull(@Storeid,Storeid)"
    sql = sql & " and  ColorID=isnull(@ColorID,ColorID)"
    sql = sql & " and  ItemSize=isnull(@ItemSize,ItemSize)"
    sql = sql & " and  ClassId=isnull(@ClassId,ClassId)"
    sql = sql & " )XTable"
    sql = sql & " group by Item_ID ,LotNO,ItemCode, ItemName"
    sql = sql & " Return"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "QryItemsInventry3", sql
 
    DB_CreateField "TblItemsUnits", "UnitWholeSalePrice", adDouble, adColNullable, , , ""
    DB_CreateField "TblSalesPricesPlanDetails", "UnitWholeSalePrice", adDouble, adColNullable, , , ""

    DB_CreateField "TblItems", "ColorID1", adInteger, adColNullable, , , ""
 
    If DB_CreateTable("TblSalesPricesPlanDetails2", True, "Id ", True) = True Then
        DB_CreateField "TblSalesPricesPlanDetails2", "PlanId", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "branch_id", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "ItemID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "UnitID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "PurchasePrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "CostPrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "SalePrice", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "UnitWholeSalePrice", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblSalesPricesPlanDetails2", "SalePriceNew", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblSalesPricesPlanDetails2", "UnitWholeSalePriceNew", adDouble, adColNullable, , , "    ", False, True
        
        db_createRelationSQL "TblSalesPricesPlan", "PlanId", "TblSalesPricesPlanDetails2", "PlanId"
        
        '   db_createRelationSQL "TblUnites", "UnitID", "TblSalesPricesPlanDetails", "UnitID"
        db_createRelationSQL "TblItems", "ItemID", "TblSalesPricesPlanDetails2", "ItemID"
  
    End If

    DB_CreateField "TblSalesPricesPlanDetails2", "SalePrice", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblOptions", "OnlyOneCashingVchr", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "CheckDateFormatCorrect", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblCustemers", "CHkMot3ahed", adBoolean, adColNullable, , , "", False, True
 
    DB_CreateField "TblCustomerContract", "CBoBasedON", adInteger, adColNullable, , , ""
    DB_CreateField "TblCustomerContract", "PlanID", adInteger, adColNullable, , , ""
                
    If DB_CreateTable("tmpAccount", True, "AdvanceID", True) = True Then
        DB_CreateField "tmpAccount", "Account_Code", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
        DB_CreateField "tmpAccount", "Account_Name", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
        DB_CreateField "tmpAccount", "Parent_Account_Code", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
        DB_CreateField "tmpAccount", "last_account", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "tmpAccount", "NewAccount_Code", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
        DB_CreateField "tmpAccount", "NewParent_Account_Code", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
        
    End If

    DB_CreateField "TblUsers", "PlaywithAuthorityMatrix", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblEmployee", "InsuranceRenew", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblEmployee", "ToM", adInteger, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblEmployee", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "CopyNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblEmployee", "NumPaspOld", adVarWChar, adColNullable, 255, , "", False, True, , True
   
    DB_CreateField "TblChangeEmployeedataDetails", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "NumPasp", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    DB_CreateField "TblCustemers", "ToPerson", adVarWChar, adColNullable, 255, , "", False, True, , True
   
    '**************************************************
   
    DB_CreateField "TbVisa", "ArriveDateH", adVarWChar, adColNullable, 20, , "      ", False, True, , True
    DB_CreateField "TbVisa", "ArriveDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "TbVisaDeti", "OfficeID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TbVisa", "OfficeID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblEmployee", "OfficeID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TbVisaDeti", "remarks", adVarWChar, adColNullable, 255, , " ", False, True, , True

    If DB_CreateTable("TblOffice", True, "id ", True) = True Then
       
        DB_CreateField "TblOffice", "Name", adVarWChar, adColNullable, 255, , " ", False, True, , True
        DB_CreateField "TblOffice", "NameE", adVarWChar, adColNullable, 255, , " ", False, True, , True
        
    End If

    DB_CreateField "TbVisaDeti", "OfficeID", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TbVisa", "OfficeID", adVarWChar, adColNullable, 255, , "", False, True, , True

    DB_CreateField "TblEmployee", "InsuranceRenew", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblEmployee", "ToM", adInteger, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblEmployee", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "CopyNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblEmployee", "NumPaspOld", adVarWChar, adColNullable, 255, , "", False, True, , True
   
    DB_CreateField "TblChangeEmployeedataDetails", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "NumPasp", adVarWChar, adColNullable, 255, , "", False, True, , True
 
    s = " ALTER VIEW emp_all_details AS"

    s = s & " SELECT dbo.TblEmpJobsTypes.JobTypeName,"
    s = s & "        dbo.TblEmpDepartments.DepartmentName,"
    s = s & "        dbo.jopstatus.color,"
    s = s & "        dbo.jopstatus.name,"
    s = s & "        dbo.TblEmployee.Emp_ID,"
    s = s & "        dbo.TblEmployee.Emp_Code,"
    s = s & "        dbo.TblEmployee.Emp_Name,"
    s = s & "        dbo.TblEmployee.Emp_Name1,"
    s = s & "        dbo.TblEmployee.Emp_Name2,"
    s = s & "        dbo.TblEmployee.Emp_Name3,"
    s = s & "        dbo.TblEmployee.Emp_Name4,"
    s = s & "        dbo.TblEmployee.Emp_Mail,"
    s = s & "        dbo.TblEmployee.Emp_Phone,"
    s = s & "        dbo.TblEmployee.Emp_mobile,"
    s = s & "        dbo.TblEmployee.Emp_Remark,"
    s = s & "        dbo.TblEmployee.Emp_Salary,"
    s = s & "        dbo.TblEmployee.Emp_Comm,"
    s = s & "        dbo.TblEmployee.EmpProfitCom,"
    s = s & "        dbo.TblEmployee.workstate,"
    s = s & "        dbo.TblEmployee.DepartmentID,"
    s = s & "        dbo.TblEmployee.JobTypeID,"
    s = s & "        dbo.TblEmployee.SpecificationID,"
    s = s & "        dbo.TblEmployee.Region,"
    s = s & "        dbo.TblEmployee.InsuranceState,"
    s = s & "        dbo.TblEmployee.InsuranceValue,"
    s = s & "        dbo.TblEmployee.OtherDiscounts,"
    s = s & "        dbo.TblEmployee.placeEkama,"
    s = s & "        dbo.TblEmployee.NumEkama,"
    s = s & "        dbo.TblEmployee.DateExpoekama,"
    s = s & "        dbo.TblEmployee.DateEndekama,"
    s = s & "        dbo.TblEmployee.DateExpoekamaH,"
    s = s & "        dbo.TblEmployee.DateEndekamah,"
    s = s & "        dbo.TblEmployee.NumLicn,"
    s = s & "        dbo.TblEmployee.DateExpLinc,"
    s = s & "        dbo.TblEmployee.DateEndLinc,"
    s = s & "        dbo.TblEmployee.DateExpLincH,"
    s = s & "        dbo.TblEmployee.DateEndLincH,"
    s = s & "        dbo.TblEmployee.NumPoket,"
    s = s & "        dbo.TblEmployee.Dateexppoket,"
    s = s & "        dbo.TblEmployee.dateendpoket,"
    s = s & "        dbo.TblEmployee.NumPasp,"
    s = s & "        dbo.TblEmployee.DateEndPasp,"
    s = s & "        dbo.TblEmployee.DateExpPasp,"
    s = s & "        dbo.TblEmployee.EmpNum,"
    s = s & "        dbo.TblEmployee.CustNum,"
    s = s & "        dbo.TblEmployee.ChekEndWork,"
    s = s & "        dbo.TblEmployee.ChekStkala,"
    s = s & "        dbo.TblEmployee.BignDateWork,"
    s = s & "        dbo.TblEmployee.EndWork,"
    s = s & "        dbo.TblEmployee.Notsstkala,"
    s = s & "        dbo.TblEmployee.checkbox1,"
    s = s & "        dbo.TblEmployee.DOB,"
    s = s & "        dbo.TblEmployee.KafelID,"
    s = s & "        dbo.TblEmployee.KafelName,"
    s = s & "        dbo.TblEmployee.pasplace,"
    s = s & "        dbo.TblEmployee.Nationality,"
    s = s & "        dbo.TblEmployee.dean,"
    s = s & "        dbo.TblEmployee.hdodno,"
    s = s & "        dbo.TblEmployee.hdoddate,"
    s = s & "        dbo.TblEmployee.hdomnfaz,"
    s = s & "        dbo.TblEmployee.kafeltel,"
    s = s & "        dbo.TblEmployee.jopstatusid,"
    s = s & "        dbo.TblEmployee.kafeladd,"
    s = s & "        dbo.TblEmployee.Emp_Salary_sakn,"
    s = s & "        dbo.TblEmployee.Emp_Salary_bus,"
    s = s & "        dbo.TblEmployee.Emp_Salary_food,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mob,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mang,"
    s = s & "        dbo.TblEmployee.Emp_Salary_others,"
    s = s & "        dbo.TblEmployee.Emp_Salary_sakn1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_bus1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_food1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_others1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mob1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mang1,"
    s = s & "        dbo.TblEmployee.Account_code,"
    s = s & "        dbo.TblEmployee.Account_code1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_saknc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_busc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_foodc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_othersc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mobc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mangc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_saknc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_busc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_foodc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_othersc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mobc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mangc1,"
    s = s & "        dbo.TblEmployee.ItemPhoto,"
    s = s & "        dbo.TblEmployee.placeWORK,"
    s = s & "        dbo.TblEmployee.project_id,"
    s = s & "        dbo.TblEmployee.Account_Code2,"
    s = s & "        dbo.TblEmployee.Dateexppoketh,"
    s = s & "        dbo.TblEmployee.dateendpoketh,"
    s = s & "        dbo.TblEmployee.opr_fullcode,"
    s = s & "        dbo.TblEmployee.term_id,"
    s = s & "        dbo.TblEmployee.opr_id,"
    s = s & "        dbo.TblEmployee.term_fullcode,"
    s = s & "        dbo.TblEmployee.BlnceVocat,"
    s = s & "        dbo.TblEmployee.InstanceDateH,"
    s = s & "        dbo.TblEmployee.InstanceDateM,"
    s = s & "        dbo.TblEmployee.PerceTage,"
    s = s & "        dbo.TblEmployee.WorkShop_Job,"
    s = s & "        dbo.TblEmployee.BYHour,"
    s = s & "        dbo.TblEmployee.Percentage,"
    s = s & "        dbo.TblEmployee.SalaryType,"
    s = s & "        dbo.TblEmployee.DriverLicenseendH,"
    s = s & "        dbo.TblEmployee.DriverLicenseStartdH,"
    s = s & "        dbo.TblEmployee.DriverLicenseend,"
    s = s & "        dbo.TblEmployee.DriverLicense,"
    s = s & "        dbo.TblEmployee.lastHolidaydateH,"
    s = s & "        dbo.TblEmployee.lastHolidaydate,"
    s = s & "        dbo.TblEmployee.OpenBalance4,"
    s = s & "        dbo.TblEmployee.OpenBalanceType4,"
    s = s & "        dbo.TblEmployee.swapedempid,"
    s = s & "        dbo.TblEmployee.mangerid,"
    s = s & "        dbo.TblEmployee.GroupID,"
    s = s & "        dbo.TblEmployee.VisaNo,"
    s = s & "        dbo.TblEmployee.JobTypeID3,"
    s = s & "        dbo.TblEmployee.JobTypeID2,"
    s = s & "        dbo.TblEmployee.JobTypeID1,"
    s = s & "        dbo.TblEmployee.LastDateH,"
    s = s & "        dbo.TblEmployee.LastDate,"
    s = s & "        dbo.TblEmployee.IssueDateH,"
    s = s & "        dbo.TblEmployee.DOBH,"
    s = s & "        dbo.TblEmployee.gradeID,"
    s = s & "        dbo.TblEmployee.InsuranceNO,"
    s = s & "        dbo.TblEmployee.BankCard,"
    s = s & "        dbo.TblEmployee.DriverId,"
    s = s & "        dbo.TblEmployee.Account_Code5,"
    s = s & "        dbo.TblEmployee.Account_Code4,"
    s = s & "        dbo.TblEmployee.Account_Code3,"
    s = s & "        dbo.TblEmployee.OpenBalanceType2,"
    s = s & "        dbo.TblEmployee.OpenBalance2,"
    s = s & "        dbo.TblEmployee.OpenBalance1,"
    s = s & "        dbo.TblEmployee.OpenBalanceType1,"
    s = s & "        dbo.TblEmployee.OpenBalance,"
    s = s & "        dbo.TblEmployee.OpenBalanceType,"
    s = s & "        dbo.TblEmployee.OpenBalanceDate,"
    s = s & "        dbo.TblEmployee.opening_balance_voucher_id,"
    s = s & "        dbo.TblEmployee.Fullcode,"
    s = s & "        dbo.TblEmployee.prifix,"
    s = s & "        dbo.TblEmployee.Emp_Namee4,"
    s = s & "        dbo.TblEmployee.Emp_Namee3,"
    s = s & "        dbo.TblEmployee.Emp_Namee2,"
    s = s & "        dbo.TblEmployee.Emp_Namee1,"
    s = s & "        dbo.TblEmployee.Emp_Namee,"
    s = s & "        dbo.TblEmployee.BranchId,"
    s = s & "        dbo.TblEmployee.cost_center_id,"
    s = s & "        dbo.jopstatus.namee,"
    s = s & "        dbo.TblEmpJobsTypes.JobTypeNamee,"
    s = s & "        dbo.TblEmpJobsTypes.VisaCode,"
    s = s & "        dbo.TblEmpDepartments.DepartmentNamee,"
    s = s & "        dbo.TblEmpDepartments.DeptColor,"
    s = s & "        dbo.TblEmpDepartments.DeptBr,"
    s = s & "        dbo.TblEmpDepartments.Dpeterial,"
    s = s & "        dbo.TblEmpDepartments.short,"
    s = s & "        dbo.EmpGroupDep.GroupName  AS LocationName,"
    s = s & "        dbo.EmpGroupDep.Fullcode   AS FullGroupCode,"
    s = s & "        dbo.EmpGroupDep.Ename      AS LocationNameE,"
    s = s & "        dbo.TblEmployee.NationlID,"
    s = s & "        InsuranceRenewA = CASE  TblEmployee.InsuranceRenew WHEN 1 THEN ' „ «· ÃœÌœ' ELSE '·„ Ì „ «· ÃœÌœ' END ,"
    s = s & "        ToMA = CASE  TblEmployee.ToM WHEN 1 THEN ' „ «· ”œÌœ' ELSE '·„ Ì „ «· ”œÌœ' END ,"
    s = s & "        TblEmployee.InsuranceRenew,"
    s = s & "        TblEmployee.ToM,"

    s = s & "        TblEmployee.InsuranceRenewDate,"
    s = s & "        TblEmployee.ToMDateNew,"
    s = s & "        TblEmployee.CopyNo,"
    s = s & "        TblEmployee.NumPaspOld"

    s = s & " From dbo.TblEmployee"
    s = s & "        LEFT OUTER JOIN dbo.EmpGroupDep"
    s = s & "             ON  dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID"
    s = s & "        LEFT OUTER JOIN dbo.jopstatus"
    s = s & "             ON  dbo.TblEmployee.jopstatusid = dbo.jopstatus.id"
    s = s & "        LEFT OUTER JOIN dbo.TblEmpDepartments"
    s = s & "             ON  dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
    s = s & "        LEFT OUTER JOIN dbo.TblEmpJobsTypes"
    s = s & "             ON  dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    Cn.Execute s

    DB_CreateField "TbVisa", "ArriveDateH", adVarWChar, adColNullable, 20, , "      ", False, True, , True
    DB_CreateField "TbVisa", "ArriveDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "TbVisaDeti", "OfficeID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TbVisa", "OfficeID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblEmployee", "OfficeID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TbVisaDeti", "remarks", adVarWChar, adColNullable, 255, , " ", False, True, , True

    If DB_CreateTable("TblOffice", True, "id ", True) = True Then
       
        DB_CreateField "TblOffice", "Name", adVarWChar, adColNullable, 255, , " ", False, True, , True
        DB_CreateField "TblOffice", "NameE", adVarWChar, adColNullable, 255, , " ", False, True, , True
        
    End If

    DB_CreateField "TbVisaDeti", "OfficeID", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TbVisa", "OfficeID", adVarWChar, adColNullable, 255, , "", False, True, , True

    DB_CreateField "TblEmployee", "InsuranceRenew", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblEmployee", "ToM", adInteger, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblEmployee", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "CopyNo", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblEmployee", "NumPaspOld", adVarWChar, adColNullable, 255, , "", False, True, , True
   
    DB_CreateField "TblChangeEmployeedataDetails", "InsuranceRenewDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "ToMDateNew", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblChangeEmployeedataDetails", "NumPasp", adVarWChar, adColNullable, 255, , "", False, True, , True
 
    s = " ALTER VIEW emp_all_details AS"

    s = s & " SELECT dbo.TblEmpJobsTypes.JobTypeName,"
    s = s & "        dbo.TblEmpDepartments.DepartmentName,"
    s = s & "        dbo.jopstatus.color,"
    s = s & "        dbo.jopstatus.name,"
    s = s & "        dbo.TblEmployee.Emp_ID,"
    s = s & "        dbo.TblEmployee.Emp_Code,"
    s = s & "        dbo.TblEmployee.Emp_Name,"
    s = s & "        dbo.TblEmployee.Emp_Name1,"
    s = s & "        dbo.TblEmployee.Emp_Name2,"
    s = s & "        dbo.TblEmployee.Emp_Name3,"
    s = s & "        dbo.TblEmployee.Emp_Name4,"
    s = s & "        dbo.TblEmployee.Emp_Mail,"
    s = s & "        dbo.TblEmployee.Emp_Phone,"
    s = s & "        dbo.TblEmployee.Emp_mobile,"
    s = s & "        dbo.TblEmployee.Emp_Remark,"
    s = s & "        dbo.TblEmployee.Emp_Salary,"
    s = s & "        dbo.TblEmployee.Emp_Comm,"
    s = s & "        dbo.TblEmployee.EmpProfitCom,"
    s = s & "        dbo.TblEmployee.workstate,"
    s = s & "        dbo.TblEmployee.DepartmentID,"
    s = s & "        dbo.TblEmployee.JobTypeID,"
    s = s & "        dbo.TblEmployee.SpecificationID,"
    s = s & "        dbo.TblEmployee.Region,"
    s = s & "        dbo.TblEmployee.InsuranceState,"
    s = s & "        dbo.TblEmployee.InsuranceValue,"
    s = s & "        dbo.TblEmployee.OtherDiscounts,"
    s = s & "        dbo.TblEmployee.placeEkama,"
    s = s & "        dbo.TblEmployee.NumEkama,"
    s = s & "        dbo.TblEmployee.DateExpoekama,"
    s = s & "        dbo.TblEmployee.DateEndekama,"
    s = s & "        dbo.TblEmployee.DateExpoekamaH,"
    s = s & "        dbo.TblEmployee.DateEndekamah,"
    s = s & "        dbo.TblEmployee.NumLicn,"
    s = s & "        dbo.TblEmployee.DateExpLinc,"
    s = s & "        dbo.TblEmployee.DateEndLinc,"
    s = s & "        dbo.TblEmployee.DateExpLincH,"
    s = s & "        dbo.TblEmployee.DateEndLincH,"
    s = s & "        dbo.TblEmployee.NumPoket,"
    s = s & "        dbo.TblEmployee.Dateexppoket,"
    s = s & "        dbo.TblEmployee.dateendpoket,"
    s = s & "        dbo.TblEmployee.NumPasp,"
    s = s & "        dbo.TblEmployee.DateEndPasp,"
    s = s & "        dbo.TblEmployee.DateExpPasp,"
    s = s & "        dbo.TblEmployee.EmpNum,"
    s = s & "        dbo.TblEmployee.CustNum,"
    s = s & "        dbo.TblEmployee.ChekEndWork,"
    s = s & "        dbo.TblEmployee.ChekStkala,"
    s = s & "        dbo.TblEmployee.BignDateWork,"
    s = s & "        dbo.TblEmployee.EndWork,"
    s = s & "        dbo.TblEmployee.Notsstkala,"
    s = s & "        dbo.TblEmployee.checkbox1,"
    s = s & "        dbo.TblEmployee.DOB,"
    s = s & "        dbo.TblEmployee.KafelID,"
    s = s & "        dbo.TblEmployee.KafelName,"
    s = s & "        dbo.TblEmployee.pasplace,"
    s = s & "        dbo.TblEmployee.Nationality,"
    s = s & "        dbo.TblEmployee.dean,"
    s = s & "        dbo.TblEmployee.hdodno,"
    s = s & "        dbo.TblEmployee.hdoddate,"
    s = s & "        dbo.TblEmployee.hdomnfaz,"
    s = s & "        dbo.TblEmployee.kafeltel,"
    s = s & "        dbo.TblEmployee.jopstatusid,"
    s = s & "        dbo.TblEmployee.kafeladd,"
    s = s & "        dbo.TblEmployee.Emp_Salary_sakn,"
    s = s & "        dbo.TblEmployee.Emp_Salary_bus,"
    s = s & "        dbo.TblEmployee.Emp_Salary_food,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mob,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mang,"
    s = s & "        dbo.TblEmployee.Emp_Salary_others,"
    s = s & "        dbo.TblEmployee.Emp_Salary_sakn1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_bus1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_food1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_others1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mob1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mang1,"
    s = s & "        dbo.TblEmployee.Account_code,"
    s = s & "        dbo.TblEmployee.Account_code1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_saknc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_busc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_foodc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_othersc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mobc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mangc,"
    s = s & "        dbo.TblEmployee.Emp_Salary_saknc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_busc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_foodc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_othersc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mobc1,"
    s = s & "        dbo.TblEmployee.Emp_Salary_mangc1,"
    s = s & "        dbo.TblEmployee.ItemPhoto,"
    s = s & "        dbo.TblEmployee.placeWORK,"
    s = s & "        dbo.TblEmployee.project_id,"
    s = s & "        dbo.TblEmployee.Account_Code2,"
    s = s & "        dbo.TblEmployee.Dateexppoketh,"
    s = s & "        dbo.TblEmployee.dateendpoketh,"
    s = s & "        dbo.TblEmployee.opr_fullcode,"
    s = s & "        dbo.TblEmployee.term_id,"
    s = s & "        dbo.TblEmployee.opr_id,"
    s = s & "        dbo.TblEmployee.term_fullcode,"
    s = s & "        dbo.TblEmployee.BlnceVocat,"
    s = s & "        dbo.TblEmployee.InstanceDateH,"
    s = s & "        dbo.TblEmployee.InstanceDateM,"
    s = s & "        dbo.TblEmployee.PerceTage,"
    s = s & "        dbo.TblEmployee.WorkShop_Job,"
    s = s & "        dbo.TblEmployee.BYHour,"
    s = s & "        dbo.TblEmployee.Percentage,"
    s = s & "        dbo.TblEmployee.SalaryType,"
    s = s & "        dbo.TblEmployee.DriverLicenseendH,"
    s = s & "        dbo.TblEmployee.DriverLicenseStartdH,"
    s = s & "        dbo.TblEmployee.DriverLicenseend,"
    s = s & "        dbo.TblEmployee.DriverLicense,"
    s = s & "        dbo.TblEmployee.lastHolidaydateH,"
    s = s & "        dbo.TblEmployee.lastHolidaydate,"
    s = s & "        dbo.TblEmployee.OpenBalance4,"
    s = s & "        dbo.TblEmployee.OpenBalanceType4,"
    s = s & "        dbo.TblEmployee.swapedempid,"
    s = s & "        dbo.TblEmployee.mangerid,"
    s = s & "        dbo.TblEmployee.GroupID,"
    s = s & "        dbo.TblEmployee.VisaNo,"
    s = s & "        dbo.TblEmployee.JobTypeID3,"
    s = s & "        dbo.TblEmployee.JobTypeID2,"
    s = s & "        dbo.TblEmployee.JobTypeID1,"
    s = s & "        dbo.TblEmployee.LastDateH,"
    s = s & "        dbo.TblEmployee.LastDate,"
    s = s & "        dbo.TblEmployee.IssueDateH,"
    s = s & "        dbo.TblEmployee.DOBH,"
    s = s & "        dbo.TblEmployee.gradeID,"
    s = s & "        dbo.TblEmployee.InsuranceNO,"
    s = s & "        dbo.TblEmployee.BankCard,"
    s = s & "        dbo.TblEmployee.DriverId,"
    s = s & "        dbo.TblEmployee.Account_Code5,"
    s = s & "        dbo.TblEmployee.Account_Code4,"
    s = s & "        dbo.TblEmployee.Account_Code3,"
    s = s & "        dbo.TblEmployee.OpenBalanceType2,"
    s = s & "        dbo.TblEmployee.OpenBalance2,"
    s = s & "        dbo.TblEmployee.OpenBalance1,"
    s = s & "        dbo.TblEmployee.OpenBalanceType1,"
    s = s & "        dbo.TblEmployee.OpenBalance,"
    s = s & "        dbo.TblEmployee.OpenBalanceType,"
    s = s & "        dbo.TblEmployee.OpenBalanceDate,"
    s = s & "        dbo.TblEmployee.opening_balance_voucher_id,"
    s = s & "        dbo.TblEmployee.Fullcode,"
    s = s & "        dbo.TblEmployee.prifix,"
    s = s & "        dbo.TblEmployee.Emp_Namee4,"
    s = s & "        dbo.TblEmployee.Emp_Namee3,"
    s = s & "        dbo.TblEmployee.Emp_Namee2,"
    s = s & "        dbo.TblEmployee.Emp_Namee1,"
    s = s & "        dbo.TblEmployee.Emp_Namee,"
    s = s & "        dbo.TblEmployee.BranchId,"
    s = s & "        dbo.TblEmployee.cost_center_id,"
    s = s & "        dbo.jopstatus.namee,"
    s = s & "        dbo.TblEmpJobsTypes.JobTypeNamee,"
    s = s & "        dbo.TblEmpJobsTypes.VisaCode,"
    s = s & "        dbo.TblEmpDepartments.DepartmentNamee,"
    s = s & "        dbo.TblEmpDepartments.DeptColor,"
    s = s & "        dbo.TblEmpDepartments.DeptBr,"
    s = s & "        dbo.TblEmpDepartments.Dpeterial,"
    s = s & "        dbo.TblEmpDepartments.short,"
    s = s & "        dbo.EmpGroupDep.GroupName  AS LocationName,"
    s = s & "        dbo.EmpGroupDep.Fullcode   AS FullGroupCode,"
    s = s & "        dbo.EmpGroupDep.Ename      AS LocationNameE,"
    s = s & "        dbo.TblEmployee.NationlID,"
    s = s & "        InsuranceRenewA = CASE  TblEmployee.InsuranceRenew WHEN 1 THEN ' „ «· ÃœÌœ' ELSE '·„ Ì „ «· ÃœÌœ' END ,"
    s = s & "        ToMA = CASE  TblEmployee.ToM WHEN 1 THEN ' „ «· ”œÌœ' ELSE '·„ Ì „ «· ”œÌœ' END ,"
    s = s & "        TblEmployee.InsuranceRenew,"
    s = s & "        TblEmployee.ToM,"

    s = s & "        TblEmployee.InsuranceRenewDate,"
    s = s & "        TblEmployee.ToMDateNew,"
    s = s & "        TblEmployee.CopyNo,"
    s = s & "        TblEmployee.NumPaspOld"

    s = s & " From dbo.TblEmployee"
    s = s & "        LEFT OUTER JOIN dbo.EmpGroupDep"
    s = s & "             ON  dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID"
    s = s & "        LEFT OUTER JOIN dbo.jopstatus"
    s = s & "             ON  dbo.TblEmployee.jopstatusid = dbo.jopstatus.id"
    s = s & "        LEFT OUTER JOIN dbo.TblEmpDepartments"
    s = s & "             ON  dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
    s = s & "        LEFT OUTER JOIN dbo.TblEmpJobsTypes"
    s = s & "             ON  dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    Cn.Execute s
 
    DB_CreateField "TblChangeEmployeedata", "IsPassport", adBoolean, adColNullable, , , "    ", False, True
    DB_CreateField "TblChangeEmployeedata", "IsInsurance", adBoolean, adColNullable, , , "    ", False, True
    DB_CreateField "TblChangeEmployeedata", "IsToM", adBoolean, adColNullable, , , "    ", False, True

    DB_CreateField "TblChangeEmployeedata", "IsToM", adBoolean, adColNullable, , , "    ", False, True

    DB_CreateField "TbVisa", "kafeladd", adVarWChar, adColNullable, 150, , "", False, True, , True
    DB_CreateField "TbVisa", "kafeltel", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TbVisa", "KafelName", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TbVisa", "KafelID", adVarWChar, adColNullable, 50, , "", False, True, , True
    
    DB_CreateField "TblItemsUnits", "OldUnitSalesPrice", adDouble, adColNullable, , , ""
    DB_CreateField "TblItemsUnits", "OldUnitWholeSalePrice", adDouble, adColNullable, , , ""
 
    DB_CreateField "TblItemsUnits", "OldUnitSalesPrice", adDouble, adColNullable, , , ""
    DB_CreateField "TblItemsUnits", "OldUnitWholeSalePrice", adDouble, adColNullable, , , ""

    DB_CreateField "TblHandWages2", "DeparmentID", adInteger, adColNullable, , , ""
    DB_CreateField "TblSalesPricesPlan", "DeparmentID", adInteger, adColNullable, , , ""


    DB_CreateField "TblCardAuthorizationReform", "IsEndAll", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemData", "NoteSerial15", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TblDefComItemData", "TransactionID5", adInteger, adColNullable, , , ""

    DB_CreateField "TblUsers", "AllowEditProductionOutManulay", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblUsers", "AllowEditVaTManulay", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transactions", "txtManulaVat", adDouble, adColNullable, 50, , " ???    ", False, True

    DB_CreateField "TblItems", "TxtBrandType", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TblItems", "TxtModel", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TblItems", "TxtColorCode", adVarWChar, adColNullable, 50, , "", False, True, , True
    DB_CreateField "TblItems", "TxtSize", adVarWChar, adColNullable, 50, , "", False, True, , True

    DB_CreateField "TblContract", "NewNO", adVarWChar, adColNullable, 255, , "", False, True, , True

    DB_CreateField "TblAqarDetai", "readyDAte", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblAqarDetai", "ready", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "notes_all", "txtManulaVat", adDouble, adColNullable, 50, , " ???    ", False, True
    DB_CreateField "Tbl_TradingContract", "txtManulaVat", adDouble, adColNullable, 50, , " ???    ", False, True

    If DB_CreateTable("TblContractReVouch", True, "ID", False) = True Then
        DB_CreateField "TblContractReVouch", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblContractReVouch", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblContractReVouch", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblContractReVouch", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblContractReVouch", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblContractReVouch", "UserID", adInteger, adColNullable, , , "  ", False, True
       
        DB_CreateField "TblContractReVouch", "AccountElecCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountElecCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountRentCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountRentCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountSai3Code", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountSai3Code2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountServiceCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountServiceCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountInsuranceCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountInsuranceCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountWaterCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountWaterCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccResPaysCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccResPaysCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
            
        DB_CreateField "TblContractReVouch", "AccInsuRefundableCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccInsuRefundableCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
           
    End If

    DB_CreateField "TblContract", "ContractReVouchID", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Notes", "ContractReVouchID", adInteger, adColNullable, , , "  ", False, True

    If DB_CreateTable("TblContractReVouch", True, "ID", False) = True Then
        DB_CreateField "TblContractReVouch", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblContractReVouch", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblContractReVouch", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblContractReVouch", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblContractReVouch", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblContractReVouch", "UserID", adInteger, adColNullable, , , "  ", False, True
       
        DB_CreateField "TblContractReVouch", "AccountElecCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountElecCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountRentCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountRentCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountSai3Code", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountSai3Code2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountServiceCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountServiceCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountInsuranceCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountInsuranceCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccountWaterCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccountWaterCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblContractReVouch", "AccResPaysCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccResPaysCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
            
        DB_CreateField "TblContractReVouch", "AccInsuRefundableCode", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblContractReVouch", "AccInsuRefundableCode2", adVarWChar, adColNullable, 100, , "      ", False, True, , True
           
    End If

    DB_CreateField "TblContractInstallments", "ContractReVouchID", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblUsers", "ShowOldAccountReports", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblEmpPassOver2", "Authorization", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblEmpPassOver2", "DeparmentID2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblEmpPassOver2", "DeptID2", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblEmpPassOver2", "KafelName", adVarWChar, adColNullable, 250, , "      ", False, True, , True
    DB_CreateField "TblContractReVouch", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "TblEmpPassOver2", "IsAuthorization", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblEmpPassOver2", "DeparmentID2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblEmpPassOver2", "DeptID2", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblEmpPassOver2", "KafelName", adVarWChar, adColNullable, 250, , "      ", False, True, , True

    DB_CreateField "TblItems", "CylinderID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "SphereID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "PackingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "UsageID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "GroupEyeID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "BaseCurveID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "ServiceID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "BreakingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "LightAdaptationID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "DIAMID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "IndexsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "CoatingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "DivisionID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblItems", "OriginID", adInteger, adColNullable, , , "", False, True

    If DB_CreateTable("tblOrigin", True, "ID", False) = True Then
        DB_CreateField "tblOrigin", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblOrigin", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblDivision", True, "ID", False) = True Then
        DB_CreateField "tblDivision", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblDivision", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblCoating", True, "ID", False) = True Then
        DB_CreateField "tblCoating", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblCoating", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblCoating", True, "ID", False) = True Then
        DB_CreateField "tblCoating", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblCoating", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblIndexs", True, "ID", False) = True Then
        DB_CreateField "tblIndexs", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblIndexs", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblDIAM", True, "ID", False) = True Then
        DB_CreateField "tblDIAM", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblDIAM", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblLightAdaptation", True, "ID", False) = True Then
        DB_CreateField "tblLightAdaptation", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblLightAdaptation", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblBreaking", True, "ID", False) = True Then
        DB_CreateField "tblBreaking", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblBreaking", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblOrigin", True, "ID", False) = True Then
        DB_CreateField "tblOrigin", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblOrigin", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblService", True, "ID", False) = True Then
        DB_CreateField "tblService", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblService", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblBaseCurve", True, "ID", False) = True Then
        DB_CreateField "tblBaseCurve", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblBaseCurve", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblGroupEye", True, "ID", False) = True Then
        DB_CreateField "tblGroupEye", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblGroupEye", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblUsage", True, "ID", False) = True Then
        DB_CreateField "tblUsage", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblUsage", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("tblPacking", True, "ID", False) = True Then
        DB_CreateField "tblPacking", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "tblPacking", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "TblEmpPassOver2", "DateCancel", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "TblItems", "MasterType", adInteger, adColNullable, , , "", False, True

    If DB_CreateTable("TblAge", True, "ID", False) = True Then
        DB_CreateField "TblAge", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblAge", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If
       
    If DB_CreateTable("TblSex", True, "ID", False) = True Then
        DB_CreateField "TblSex", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblSex", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "Groups", "AqrCompenetId", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblAqrCompenet", "GroupId", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblUsers", "CanOpenWorkOrder", adBoolean, adColNullable, , , "", False, True
    

 
    DB_CreateField "Transactions", "dbname", adVarWChar, adColNullable, 250, , "      ", False, True, , True
    DB_CreateField "Transactions", "ServerName", adVarWChar, adColNullable, 250, , "      ", False, True, , True

    DB_CreateField "Transactions", "Iqar", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Transactions", "UnitType", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Transactions", "UnitNo", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "Groups", "dbname", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblAqrCompenet", "GroupId", adInteger, adColNullable, , , "  ", False, True
       
    DB_CreateField "Transactions", "dbname", adVarWChar, adColNullable, 250, , "      ", False, True, , True
    DB_CreateField "Transactions", "ServerName", adVarWChar, adColNullable, 250, , "      ", False, True, , True

    DB_CreateField "Transactions", "Iqar", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Transactions", "UnitType", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Transactions", "UnitNo", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "Groups", "IsShowCover", adBoolean, adColNullable, , , "  ", False, True

    DB_updateField "Transactions", "UnitNo", "nvarchar(255)   "

    DB_CreateField "TblPripaidExpensesDet", "CostCenterID", adVarWChar, adColNullable, 250, , "      ", False, True, , True
    DB_CreateField "TblPripaidExpensesDet", "CostCenterIDName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

    DB_CreateField "TblPaytAmortizationDet", "CostCenterID", adVarWChar, adColNullable, 250, , "      ", False, True, , True
    DB_CreateField "TblPaytAmortizationDet", "CostCenterIDName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
    DB_CreateField "TblPaytAmortizationDet", "LineNo1", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
    DB_updateField "Transactions", "UnitNo", "nvarchar(255)   "

    'DB_CreateField "Groups", "AqrCompenetID", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "TblPripaidExpensesDet", "ProjectID", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblPripaidExpensesDet", "ProjectName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

    DB_CreateField "TblPaytAmortizationDet", "ProjectID", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "TblPaytAmortizationDet", "ProjectName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

    DB_CreateField "TblAqarDetai", "Mremarks", adVarWChar, adColNullable, 4000, , "", False, True, , True

    DB_CreateField "tblItems", "IsNotShowAlarm", adBoolean, adColNullable, , , "  ", False, True
    DB_CreateField "TblContract", "IsNotCreateEntry", adInteger, adColNullable, , , " ???    ", False, True
    DB_CreateField "TblOptions", "CantRepetttransferNoforCashing", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblFiterWaiver", "UnitElectric", adVarWChar, adColNullable, 2500, , "      ", False, True, , True
    DB_CreateField "TblFiterWaiver", "NewNO", adVarWChar, adColNullable, 2500, , "      ", False, True, , True
    DB_CreateField "TblFiterWaiver", "Accredit", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblFiterWaiver", "ManualNO", adVarWChar, adColNullable, 50, , "      ", False, True, , True

    DB_CreateField "TblCarsData", "DCOwner", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
    DB_CreateField "TblCarsData", "ChkOtherVendor", adBoolean, adColNullable, , , "  ", False, True

    DB_CreateField "notes_all", "TxtRent", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblOrderUpload", "TxtRent", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTravDueKDet", "TxtRent", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblOptions", "CheckMobileFormatCorrect", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "ExpensesType", "DataTypeExchangeCode", adInteger, adColNullable, , , "  ", False, True

    '*******************************************************************************************************

'    UpdateDataBasePart27
 
End Function
Function UpdateDataBasePart27()

    On Error Resume Next
    Dim New_View As String
    Dim s        As String
    
    '*************************
    
    UpdateDataBasePart30
    
    sql = "    DROP FUNCTION [QryLastItemsPurPrice]" & CHR(13)
    Cn.Execute sql






    '  sql = " CREATE FUNCTION QryItemsTransactionsTotalsByStores(@TransType int =0,@TransType2 int=0,@TransType3 int=0,@FromDate datetime ,@ToDate datetime ,@storeid as integer,@ItemID  as integer,@Transaction_ID as float=null )" & CHR(13)
    sql = "CREATE FUNCTION [dbo].[QryLastItemsPurPrice] (@X int,@Y int,@ToDate DateTime='31-DEC-9999')   " & CHR(13)
    sql = sql & "  RETURNS @XTable Table" & CHR(13)
    sql = sql & "   (" & CHR(13)
    sql = sql & "      LastPurTrans    int," & CHR(13)
    sql = sql & "      Transaction_ID  int," & CHR(13)
    sql = sql & "      Transaction_Serial nvarchar(50)," & CHR(13)
    sql = sql & "      Transaction_Date smalldatetime," & CHR(13)
    sql = sql & "      ItemSerial  nvarchar(50)," & CHR(13)
    sql = sql & "      Price money," & CHR(13)
    sql = sql & "      ItemName nvarchar(255)," & CHR(13)
    sql = sql & "      ItemCode nvarchar(50)," & CHR(13)
    sql = sql & "      ItemID int," & CHR(13)
    sql = sql & "      CusName nvarchar(50)" & CHR(13)
    sql = sql & "   )" & CHR(13)
    sql = sql & "   AS" & CHR(13)
    sql = sql & "    Begin" & CHR(13)
    sql = sql & "    INSERT  @XTable" & CHR(13)
    sql = sql & "   SELECT     * From" & CHR(13)
    sql = sql & "  (" & CHR(13)
    sql = sql & "  SELECT QryLastPurItemsTrans.LastPurTrans,dbo.Transactions.Transaction_ID,dbo.Transactions.Transaction_Serial," & CHR(13)
    sql = sql & "   dbo.Transactions. Transaction_Date,dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Price,dbo.TblItems.ItemName," & CHR(13)
    sql = sql & "    dbo.TblItems.ItemCode,QryLastPurItemsTrans.ItemID , dbo.TblCustemers.CusName  FROM dbo.Transactions INNER JOIN" & CHR(13)
    sql = sql & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID  INNER JOIN" & CHR(13)
    sql = sql & "   dbo.QryLastPurItemsTrans(@X,@Y,@ToDate) QryLastPurItemsTrans INNER JOIN dbo.TblItems ON" & CHR(13)
    sql = sql & "  QryLastPurItemsTrans.ItemID = dbo.TblItems.ItemID ON  dbo.Transactions.Transaction_ID = QryLastPurItemsTrans.LastPurTrans" & CHR(13)
    sql = sql & "   left  JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID" & CHR(13)
    sql = sql & "   Where ((dbo.Transaction_Details.Item_ID = QryLastPurItemsTrans.ItemID)" & CHR(13)
    sql = sql & "  and Transactions.Transaction_Date=@ToDate)" & CHR(13)
    sql = sql & "  )DERIVEDTBL" & CHR(13)
    sql = sql & "    ORDER BY LastPurTrans   Return  End" & CHR(13)

    db_createOrUpdateFuctionSQL "QryLastItemsPurPrice", sql

    '*******************************************************************************************************
    DB_CreateField "TblItems", "lowering2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblItems", "ItemLimit", adDouble, adColNullable, , , "    ", False, True
    
   DB_CreateField "Transaction_Details", "Account_Code", adVarWChar, adColNullable, 50, , "      ", False, True, , True
   DB_CreateField "project_bill_details", "AccountCode", adVarWChar, adColNullable, 50, , "      ", False, True, , True
   DB_CreateField "SubcontractorContract2", "AccountCode", adVarWChar, adColNullable, 50, , "      ", False, True, , True
   DB_CreateField "project_billl", "DiscountAccount", adVarWChar, adColNullable, 50, , "      ", False, True, , True
    
    DB_CreateField "TBLLC", "PercentV", adDouble, adColNullable, , , "    ", False, True
    
    
    DB_CreateField "TblContract", "InsuranceValue1", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContract", "InsuranceValueAdd", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "project_billl", "DiscountGMater", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "FixedAssets", "Quantity", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblPaymentType", "TaxTobacco", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Transactions", "TaxTobacco", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblPaymentType", "AccTaxTobacco", adVarWChar, adColNullable, 50, , "      ", False, True, , True
    DB_CreateField "Transactions", "IsHiddenVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Transactions", "SerPos", adInteger, adColNullable, , , "        ", False, True
    DB_CreateField "TblPaymentType", "IsNewCode", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblPaymentType", "IsHiddenVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblPaymentType", "IsDefault", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblContractReVouch", "IsVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblContractInstallments", "ContractReVouchID2", adInteger, adColNullable, , , "        ", False, True

    DB_CreateField "TblContractReVouch", "IsVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblContractInstallments", "ContractReVouchID2", adInteger, adColNullable, , , "        ", False, True

    DB_CreateField "TblContractInstallments", "VATValueOld", adDouble, adColNullable, , , "        ", False, True
    '************************
    
    
    DB_CreateField "TblProductOrderFactoryExpenses", "Vat", adDouble, adColNullable, , , "        ", False, True
    DB_CreateField "TblProductOrderFactoryExpenses", "PriceTotal", adDouble, adColNullable, , , "        ", False, True
    DB_CreateField "TblProductOrderFactoryExpenses", "Vatyo", adDouble, adColNullable, , , "        ", False, True
    DB_CreateField "TblProductOrderFactoryExpenses", "FlgVat", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblProductOrderFactoryExpenses", "ForcedFlg", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblProductOrderFactoryExpenses", "CurrRow", adDouble, adColNullable, , , "        ", False, True
               
               
               
    
    
    UpdateCostriceProcedure
    
 '   UpdateCostriceProcedureByStores

    '**************************
    If DB_CreateTable("TblSalesPricesPlanDetails3", True, "[PlaneId] [INT] NOT NULL,[Ser] [INT] NOT NULL", False, "[PlaneId] ASC,[Ser] ASC") = True Then
        DB_CreateField "TblSalesPricesPlanDetails3", "FromPrice", adCurrency, adColNullable, , , "", False, True
        DB_CreateField "TblSalesPricesPlanDetails3", "ToPrice", adCurrency, adColNullable, , , "", False, True
        DB_CreateField "TblSalesPricesPlanDetails3", "Result", adCurrency, adColNullable, , , "", False, True
        DB_CreateField "TblSalesPricesPlanDetails3", "Example", adCurrency, adColNullable, , , "", False, True
    End If

    '**************************

    add_record_to_table "TransactionTypes", "Transaction_Type,TransactionTypeName,TransactionEnglishName,StockEffect", " 76 , ' «·€«¡ «·ÕÃ“' , 'Cancellation of reservation' ,1", "Transaction_Type", 76
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "IsHidden", adInteger, adColNullable, , , "      ", False, True

    s = " SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,                      dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,                       dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,                      dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,                      dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,                      dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, "
    s = s & " dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID , dbo.transactions.Transaction_serial, dbo.transactions.Transaction_Date ,"
    s = s & " dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.Accounts.account_serial, dbo.Accounts.Account_NameEng, dbo.Accounts.Parent_Account_Code, dbo.Accounts.opening_balance, dbo.Accounts.opening_balance_type, dbo.Accounts.Branch, dbo.Accounts.Sum_account, dbo.Accounts.cost_center, dbo.Accounts.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.TblNotesTypes.NotesTypeNameE, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.ActivityTypeId, "
    s = s & " dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,                      "
    s = s & " dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency,                      dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee,                      dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id,                      dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid,                      dbo.DOUBLE_ENTREY_VOUCHERS.pandid , dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo     FROM         dbo.TblAqar RIGHT OUTER JOIN                      dbo.TblBranchesData INNER JOIN                      dbo.TblUsers INNER JOIN                      dbo.DOUBLE_ENTREY_VOUCHERS ON "
    s = s & " dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON                      dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON                      dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN                      dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN                      dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN                      dbo.Notes LEFT OUTER JOIN                      dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN                      dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type     "
    s = s & " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null and IsNull(DOUBLE_ENTREY_VOUCHERS.IsHidden,0) =0)"

    db_createOrUpdateviewSQL "RptLedger_Sub", s

    s = " Create FUNCTION GetInstalDiscuValueByDate (@FixedID  integer ,@ToDate datetime )"
    s = s & "   RETURNS Float    AS    Begin"
    s = s & "     RETURN ("
    s = s & "       SELECT"
    s = s & "       SumVal =Sum("
    s = s & "      CASE FAVType WHEN 1 THEN"
    s = s & "         (ExcludedValuePrt)"
    s = s & "         Else"
    s = s & "       currentvalue"
    s = s & "       END)"
       
    s = s & "    From dbo.notes_all"
    s = s & "      WHERE     (NoteType = 8028) AND (FAID = @FixedID) AND (NoteDate <= @ToDate)"
    s = s & "       GROUP BY FAID"
    s = s & "     )"
    s = s & "   End"

    s = s & " "

    db_createOrUpdateFuctionSQL "GetInstalDiscuValueByDate", s
    DB_CreateField "TblContract", "IsShamel", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "TblyearsData", "IsFirstYear", adBoolean, adColNullable, , , "        ", False, True

    DB_CreateField "notes_all", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "notes_all", "FromDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
    DB_CreateField "notes_all", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "notes_all", "ToDateH", adVarWChar, adColNullable, 50, , "      ", False, True, , True
    'notes_all.FromDate,notes_all.FromDateH,notes_all.ToDate,notes_all.ToDateH,"
    DB_CreateField "Transactions", "CardId0", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "Transactions", "CardId1", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "Transactions", "chkIsFirstInv", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblEmployee", "chkShowTasks", adBoolean, adColNullable, , , "    ", False, True
    DB_CreateField "tblusers", "CanEditMinRentValue", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "tblGeneralCashingDetails", "Returnvalue", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblAging", "TypeTrans", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblAging", "IsDeleted", adInteger, adColNullable, , , "    ", False, True
   DB_CreateField "TblAging", "ItemId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "branches", "a214", adVarWChar, adColNullable, 90, , "", False, True, , True
    DB_CreateField "TblEmployee", "Commission", adCurrency, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsMashghal", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "CustCreat4Acc", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "SuppCreat4Acc", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "CreateEntryBillItems", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsCheque", adBoolean, adColNullable, , , "", False, True
      


    DB_CreateField "tblOPtions", "IsSalesOrder", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsQrCodePrint", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsShowItemsBranch", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsElecWaterCont", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsDogeMode", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "tblOPtions", "IsMaintItemMode", adBoolean, adColNullable, , , "", False, True
    
    
    DB_CreateField "tblOPtions", "IsHeaderPrint", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblCardAuthorizationReform", "ItemID33", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblCardAuthorizationReform", "SalesInvoiceOrder", adVarWChar, adColNullable, 250, , "", False, True, , True

    DB_CreateField "TblCardAuthorizationReform", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblCardAuthorizationReform", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblCardAuthorizationReform", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
   
    DB_CreateField "TblCarBillMentains", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblCarBillMentains", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblCarBillMentains", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
  
    DB_CreateField "TblHandWages", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblHandWages", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblHandWages", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "branches", "a790", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a791", adVarWChar, adColNullable, 250, , "", False, True, , True

    DB_CreateField "branches", "a212", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a213", adVarWChar, adColNullable, 250, , "", False, True, , True
    
    DB_CreateField "branches", "a217", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a215", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a216", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a217", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a218", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a219", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a220", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a221", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a222", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a223", adVarWChar, adColNullable, 250, , "", False, True, , True
    DB_CreateField "branches", "a224", adVarWChar, adColNullable, 250, , "", False, True, , True


    DB_CreateField "tblOPtions", "Isthickness", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblJobOrders2", "ItemID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblJobOrders2", "Emp_ID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblJobOrders2", "RemarkItem", adVarWChar, adColNullable, 2500, , "      ", False

    DB_CreateField "TblTravDueK", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblTravDueK", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblTravDueK", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "project_billl", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "project_billl", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "project_billl", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblContractInstallments", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblContractInstallments", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "TblContract", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblContract", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblContract", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "Notes", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Notes", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Notes", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "Notes_All", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Notes_All", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Notes_All", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "Notes_All", "TotalFines", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Notes_All", "RequestNo", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Notes_All", "ContractNo", adVarWChar, adColNullable, 255, , "      ", False
      
    DB_CreateField "tblOPtions", "CountPrint", adVarWChar, adColNullable, 10, , "      ", False

    DB_CreateField "TblDefComItem", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblDefComItem", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "TblDefComItem", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "Transactions", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "Transactions", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    '
    'DB_CreateField "TblEmpData", "Photo2", adLongVarBinary, adColNullable, , , " Â·     ⁄„· »«·»—«ﬂÊœ «·«’‰«› ", False, True
    DB_CreateField "Transactions", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True

    DB_CreateField "tblOPtions", "Company_QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "tblItems", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblCustemers", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblCountriesData", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblTypeVats", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblTypeActivity", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "tblUnites", "QRCODE", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblCountriesData", "ECountryName", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True

    DB_CreateField "Subject_doc", "IsDeleted", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblUnites", "HaveWeight", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItem", "CBoBasedON", adInteger, adColNullable, , , "    ", False, True
    
    DB_CreateField "tmpPos33", "Transaction_ID", adInteger, adColNullable, , , ""

     
    If DB_CreateTable("tmpPos33", True, "ID", True) = False Then
            DB_CreateField "tmpPos33", "NetValue0", adDouble, adColNullable, , , "    ", False, True
    End If
        DB_CreateField "tmpPos33", "Transaction_NetValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "TotalNetValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "NetValue0", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "NetValue1", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "NetValue2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "NetValue3", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tmpPos33", "NetValue4", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "tmpPos33", "NoteSerial1", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        
        
        DB_CreateField "tmpPos33", "ID0", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "ID1", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "ID2", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "ID3", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "ID4", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "boxid", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "CurrentCashireID", adInteger, adColNullable, , , ""
        DB_CreateField "tmpPos33", "PointID", adInteger, adColNullable, , , ""
        
    If DB_CreateTable("TblSphCylCusType", True, "ID", False) = True Then
        DB_CreateField "TblSphCylCusType", "Name", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblSphCylCusType", "AGE", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblSphCylCusType", "CusID", adInteger, adColNullable, , , ""
    End If

    If DB_CreateTable("TblSphCylCusType2", True, "ID", False) = True Then
        DB_CreateField "TblSphCylCusType2", "MasterID", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "SerID", adInteger, adColNullable, , , ""
            
        DB_CreateField "TblSphCylCusType2", "SphereIDR", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "SphereIDL", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "CylinderIDR", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "CylinderIDL", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "SphereIDL", adInteger, adColNullable, , , ""
        DB_CreateField "TblSphCylCusType2", "CYLR", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblSphCylCusType2", "CYLL", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblSphCylCusType2", "SPHL", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblSphCylCusType2", "SPHR", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
            
        DB_CreateField "TblSphCylCusType2", "AxisR", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblSphCylCusType2", "AxisL", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
            
        DB_CreateField "TblSphCylCusType2", "IPD", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
            
    End If

    DB_CreateField "TblSphCylCusType2", "CYLL", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblSphCylCusType2", "CylinderIDL", adInteger, adColNullable, , , ""

    If DB_CreateTable("TblTamimi", True, "ID", False) = True Then

        DB_CreateField "TblTamimi", "NetSalesAfter1", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter3", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter4", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter5", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter6", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter7", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter8", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter9", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi", "NetSalesAfter10", adDouble, adColNullable, , , "    ", False, True
                
        DB_CreateField "TblTamimi", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTamimi", "FromDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTamimi", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTamimi", "CusID", adInteger, adColNullable, , , ""
    End If

    DB_CreateField "TblTamimi", "XPTxtID1", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "XPTxtID2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "XPTxtID3", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "XPTxtID4", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "XPTxtID5", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "tblOPtions", "IsBlue", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblTamimi", "NetSalesAfter11", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "NetSalesAfter12", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblTamimi", "NetSalesAfter13", adDouble, adColNullable, , , "    ", False, True

    If DB_CreateTable("TblTamimi2", True, "ID", True) = True Then
        DB_CreateField "TblTamimi2", "RECEIVINGDATE", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblTamimi2", "Value", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi2", "[Percent]", adDouble, adColNullable, , , "    ", False, True
                
        DB_CreateField "TblTamimi2", "TAMIMIDOCNO", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
        DB_CreateField "TblTamimi2", "GROSSAMOUNT", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi2", "DISCOUNT", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi2", "NETAMOUNT", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi2", "DISTRIBUTION", adDouble, adColNullable, , , "    ", False, True
                
        DB_CreateField "TblTamimi2", "TypeN2", adInteger, adColNullable, , , ""
        DB_CreateField "TblTamimi2", "TypeN", adInteger, adColNullable, , , ""
    End If

    DB_CreateField "TblTamimi2", "MasterID", adInteger, adColNullable, , , ""
                
    DB_CreateField "TblTamimi3", "MasterID", adInteger, adColNullable, , , ""
    DB_CreateField "TblTamimi2", "INVOICENUMBER", adVarWChar, adColNullable, 3000, , "      ", False, True, , True

    If DB_CreateTable("TblTamimi3", True, "ID", True) = True Then

        DB_CreateField "TblTamimi3", "Value", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblTamimi3", "Percent", adDouble, adColNullable, , , "    ", False, True
                
        DB_CreateField "TblTamimi3", "Account_code", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
        DB_CreateField "TblTamimi3", "AccountName", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
        DB_CreateField "TblTamimi3", "TypeN2", adInteger, adColNullable, , , ""
    End If

    DB_CreateField "TblTamimi3", "Percent", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblDefComItemData", "Diameter2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "thickness2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "widtj2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "DO2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "DI2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "hight2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItemData", "Length2", adDouble, adColNullable, , , "    ", False, True
        
    DB_CreateField "Transactions", "LblDiscountsTotal", adDouble, adColNullable, , , "    ", False, True
        
    DB_CreateField "TblRegDateDelgate", "VisitTime0", adVarWChar, adColNullable, 20, , "      ", False, True, , True
    DB_CreateField "TblRegDateDelgate", "VisitTime1", adVarWChar, adColNullable, 20, , "      ", False, True, , True
        
    DB_CreateField "TblRegDateDelgate", "DateVis0", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblRegDateDelgate", "DateVis1", adDBTimeStamp, adColNullable, , , "      ", False, True
        
    DB_CreateField "TblRegDateDelgate", "GPS0", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblRegDateDelgate", "GPS1", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblRegDateDelgate", "Address0", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblRegDateDelgate", "Address1", adVarWChar, adColNullable, 3000, , "      ", False, True, , True

    DB_CreateField "SubcontractorContract", "discount1value", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "SubcontractorContract", "discount2value", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "SubcontractorContract", "discount3value", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "SubcontractorContract", "discount1ID", adInteger, adColNullable, , , ""
    DB_CreateField "SubcontractorContract", "discount2ID", adInteger, adColNullable, , , ""
    DB_CreateField "SubcontractorContract", "subContractorId", adInteger, adColNullable, , , ""

    DB_CreateField "TblCustomerContract", "IsLastMonth", adBoolean, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustomerContract", "Percent1", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustomerContract", "Percent2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustomerContract", "Percent3", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustomerContract", "Percent4", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblCustomerContract", "AccCode1", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblCustomerContract", "AccCode2", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblCustomerContract", "AccCode3", adVarWChar, adColNullable, 3000, , "      ", False, True, , True
    DB_CreateField "TblCustomerContract", "AccCode4", adVarWChar, adColNullable, 3000, , "      ", False, True, , True

    DB_CreateField "transactions", "ManualDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "transactions", "RegDate", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "TblOffline", "CountSalesOfeers", adInteger, adColNullable, , , ""
    DB_CreateField "TblOffline", "CountDefComItem", adInteger, adColNullable, , , ""

    DB_CreateField "TblDefComItem", "DepandToConv", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblDefComItem", "SessionCode", adVarWChar, adColNullable, 255, , "ECEE    ", False, True, , True
    DB_CreateField "TblDefComItem", "Copied", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblDefComItem", "OldID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblUnites", "HaveWeight", adInteger, adColNullable, , , "    ", False, True
    Dim j As Integer

    For j = 0 To 6
        DB_CreateField "TblVATAvowal", "txtExpenses" & j, adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblVATAvowal", "txtExpensesVat" & j, adDouble, adColNullable, , , "    ", False, True
    Next j

    DB_CreateField "Transaction_Details", "StoreIDLoc", adInteger, adColNullable, , , " ??? ", False, True
    DB_CreateField "Transaction_Details", "OldID", adInteger, adColNullable, , , " ??? ", False, True

    If DB_CreateTable("TblTypeFarm", True, "ID", False) = True Then
        DB_CreateField "TblTypeFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If
   
    If DB_CreateTable("Translations", True, "ID", True) = True Then
        DB_CreateField "Translations", "OldArabic", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations", "Arabic", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations", "English", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations", "RowId", adInteger, adColNullable, , , " ??? ", False, True
        DB_CreateField "Translations", "IsVisible", adBoolean, adColNullable, , , " ??? ", False, True
        DB_CreateField "Translations", "ControlName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations", "ControlIndex", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations", "FormName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
 
    End If

    If DB_CreateTable("Translations2", True, "ID", True) = True Then
        DB_CreateField "Translations2", "OldArabic", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations2", "Arabic", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations2", "English", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations2", "RowId", adInteger, adColNullable, , , " ??? ", False, True
        DB_CreateField "Translations2", "IsVisible", adBoolean, adColNullable, , , " ??? ", False, True
        DB_CreateField "Translations2", "ControlName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations2", "ControlIndex", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "Translations2", "FormName", adVarWChar, adColNullable, 400, , "      ", False, True, , True

    End If

    If DB_CreateTable("TblTypeImage", True, "ID", False) = True Then
        DB_CreateField "TblTypeImage", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeImage", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblTypeImage2", True, "ID", False) = True Then
        DB_CreateField "TblTypeImage2", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeImage2", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeImage2", "MasterId", adInteger, adColNullable, , , "      ", False, True
    End If

    If DB_CreateTable("TblTypeActivity", True, "ID", False) = True Then
        DB_CreateField "TblTypeActivity", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeActivity", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeActivity", "MasterId", adInteger, adColNullable, , , "      ", False, True
    End If

    If DB_CreateTable("TblTypeVats", True, "ID", False) = True Then
        DB_CreateField "TblTypeVats", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeVats", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblTypeVats", "MasterId", adInteger, adColNullable, , , "      ", False, True
    End If

    DB_CreateField "subjects_images", "Type1", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "subjects_images", "Type2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "subjects_images", "Type1ID", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "subjects_images", "Type2ID", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "subjects_images", "ContNo", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "subjects_images", "NoteSerial", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "subjects_images", "Datee1", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "subjects_images", "Datee2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "subjects_images", "NameFile", adVarWChar, adColNullable, 4000, , "", False, True, , True

    DB_CreateField "Subject_doc", "Type1", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Subject_doc", "Type2", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Subject_doc", "Type1ID", adInteger, adColNullable, , , "  ", False, True
    DB_CreateField "Subject_doc", "Type2ID", adInteger, adColNullable, , , "  ", False, True

    DB_CreateField "Subject_doc", "ContNo", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "Subject_doc", "NoteSerial", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "Subject_doc", "Datee1", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "Subject_doc", "Datee2", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "Subject_doc", "NameFile", adVarWChar, adColNullable, 4000, , "", False, True, , True

    DB_CreateField "Transaction_Details", "IsLastPurPrice", adInteger, adColNullable, , , " ??? ", False, True

    DB_CreateField "TblOptions", "AllowItemByRowMove", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblOptions", "AllowItemByRowOut", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblItems", "ItemWithOutVAT", adInteger, adColNullable, , , "    ", False, True

    If DB_CreateTable("TblCategoryFarm", True, "ID", False) = True Then
        DB_CreateField "TblCategoryFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCategoryFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "TblCustemers", "CUsDOB2", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
    DB_CreateField "TblCustemers", "NationalNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True

    If DB_CreateTable("TblCustResp", True, "ID", True) = True Then
        DB_CreateField "TblCustResp", "CusID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblCustResp", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp", "NationalNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp", "Tel1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp", "Tel2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp", "EmailW", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp", "jobName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
           
    End If

    DB_CreateField "TblCustemers", "txtGeneralTax", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtTaxC", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtTaxStamp", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtWorkEarningTaxes", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtTaxNo1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    



DB_CreateField "TblCustemers", "ParentAccountCurrentAss", adVarWChar, adColNullable, 400, , "      ", False, True, , True
DB_CreateField "TblCustemers", "ParentAccountCurrentHih", adVarWChar, adColNullable, 400, , "      ", False, True, , True
DB_CreateField "TblCustemers", "Account_CodeAss1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
DB_CreateField "TblCustemers", "Account_CodeAss2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
DB_CreateField "TblCustemers", "Account_CodeHi1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
DB_CreateField "TblCustemers", "Account_CodeHi2", adVarWChar, adColNullable, 400, , "      ", False, True, , True

    
    DB_CreateField "TblCustemers", "txtTaxNo2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtTaxNo3", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtTaxNo4", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            
    DB_CreateField "TblCustemers", "txtTaxNo5", adVarWChar, adColNullable, 400, , "      ", False, True, , True
             
    DB_CreateField "TblCustemers", "TxtVATNO1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "TxtVATNO2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "TxtVATNO3", adVarWChar, adColNullable, 400, , "      ", False, True, , True
             
    DB_CreateField "TblCustemers", "txtInsOffice", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtCardImportNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtCardExportNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtGeneralTax", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    DB_CreateField "TblCustemers", "txtGeneralTax", adVarWChar, adColNullable, 400, , "      ", False, True, , True
           
    DB_CreateField "TblCustemers", "txtDateIssuanceImport", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCustemers", "txtDateRenewImport", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCustemers", "txtDateIssuanceExport", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCustemers", "txtDateRenewExport", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCustemers", "txtLastDateRenewReg", adDBTimeStamp, adColNullable, , , "      ", False, True
           
    DB_CreateField "TblCustemers", "Period1", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustemers", "Period2", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblCustemers", "cmbTaxStatus", adInteger, adColNullable, , , "    ", False, True
           
    If DB_CreateTable("TblCustResp2", True, "ID", True) = True Then
        DB_CreateField "TblCustResp2", "CusID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblCustResp2", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "NationalNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "Tel1", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "Tel2", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "EmailW", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "jobName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblCustResp2", "DOB", adDBTimeStamp, adColNullable, , , "      ", False, True
           
    End If

    DB_CreateField "TblCustResp2", "DOB", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblCustResp", "DOB", adDBTimeStamp, adColNullable, , , "      ", False, True

    If DB_CreateTable("TblStrain", True, "ID", False) = True Then
        DB_CreateField "TblStrain", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblStrain", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblAdjectiveFarm", True, "ID", False) = True Then
        DB_CreateField "TblAdjectiveFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblAdjectiveFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblOwnersFarm", True, "ID", False) = True Then
        DB_CreateField "TblOwnersFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblOwnersFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblLocationFarm", True, "ID", False) = True Then
        DB_CreateField "TblLocationFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblLocationFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblStatusFarm", True, "ID", False) = True Then
        DB_CreateField "TblStatusFarm", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblStatusFarm", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "tblContractInsAllocationsDetails", "ContNo", adInteger, adColNullable, , , "  ", False, True

    If DB_CreateTable("TblFarmAnimalRegister", True, "ID", False) = True Then
        DB_CreateField "TblFarmAnimalRegister", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmAnimalRegister", "StatusFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister", "CurrentWeight", adDouble, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalRegister", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmAnimalRegister", "DateBirth", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmAnimalRegister", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblFarmAnimalRegister", "TypeFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister", "CategoryFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister", "StrainId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister", "AdjectiveFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister", "OwnersFarmID", adInteger, adColNullable, , , "  ", False, True
      
        DB_CreateField "TblFarmAnimalRegister", "LocationFarmID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalRegister", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalRegister", "UserID", adInteger, adColNullable, , , "  ", False, True
       
        DB_CreateField "TblFarmAnimalRegister", "Name", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        DB_CreateField "TblFarmAnimalRegister", "FatherName", adVarWChar, adColNullable, 100, , "      ", False, True, , True
        
        DB_CreateField "TblFarmAnimalRegister", "MotherName", adVarWChar, adColNullable, 100, , "      ", False, True, , True
           
    End If
         
    If DB_CreateTable("TblFarmAnimalRegister2", True, "ID", True) = True Then
                
        DB_CreateField "TblFarmAnimalRegister2", "SerID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFarmAnimalRegister2", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
                
        DB_CreateField "TblFarmAnimalRegister2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmAnimalRegister2", "StatusFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister2", "CurrentWeight", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalRegister2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    End If

    If DB_CreateTable("TblFarmRequestTreatment", True, "ID", False) = True Then
        DB_CreateField "TblFarmRequestTreatment", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmRequestTreatment", "Discription", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmRequestTreatment", "Discription2", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmRequestTreatment", "DoctorName", adVarWChar, adColNullable, 255, , "      ", False
        
        DB_CreateField "TblFarmRequestTreatment", "CurrentWeight", adDouble, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmRequestTreatment", "DateBirth", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmRequestTreatment", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "FarmAnimalRegisterId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "TypeFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "CategoryFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "StrainId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "AdjectiveFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "OwnersFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "LocationFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment", "StatusFarmID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmRequestTreatment", "UserID", adInteger, adColNullable, , , "  ", False, True
           
    End If
          
    If DB_CreateTable("TblFarmRequestTreatment2", True, "ID", True) = True Then
                
        DB_CreateField "TblFarmRequestTreatment2", "SerID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "TblFarmRequestTreatment2", "MasterID", adInteger, adColNullable, , , " ???    ", False, True
                
        DB_CreateField "TblFarmRequestTreatment2", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmRequestTreatment2", "StatusFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment2", "CurrentWeight", adDouble, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmRequestTreatment2", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    End If
     
    If DB_CreateTable("TblFarmAnimalDeath", True, "ID", False) = True Then
        DB_CreateField "TblFarmAnimalDeath", "Remarks", adVarWChar, adColNullable, 255, , "      ", False
        DB_CreateField "TblFarmAnimalDeath", "FarmAnimalRegisterId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmAnimalDeath", "DateBirth", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmAnimalDeath", "ToDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblFarmAnimalDeath", "SexID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalDeath", "TypeFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "CategoryFarmId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "StrainId", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "AdjectiveFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "OwnersFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "LocationFarmID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblFarmAnimalDeath", "StatusFarmID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalDeath", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalDeath", "BranchId", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblFarmAnimalDeath", "UserID", adInteger, adColNullable, , , "  ", False, True
           
    End If
     
    DB_CreateField "TblUsers", "AllowSelectEmp", adBoolean, adColNullable, , , "    ", False, True

    DB_CreateField "TblFarmAnimalRegister", "SexID", adInteger, adColNullable, , , "  ", False, True

    If DB_CreateTable("TblFarmDoctors", True, "ID", False) = True Then
        DB_CreateField "TblFarmDoctors", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblFarmDoctors", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    If DB_CreateTable("TblInvoiceError", True, "ID", True) = True Then
        DB_CreateField "TblInvoiceError", "Transaction_Type", adInteger, adColNullable, , , " ??? ", False, True
        DB_CreateField "TblInvoiceError", "Transaction_ID", adInteger, adColNullable, , , " ??? ", False, True
        DB_CreateField "TblInvoiceError", "NoteSerial1", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
        DB_CreateField "TblInvoiceError", "Notes", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
    End If

    DB_CreateField "TblFarmRequestTreatment", "DoctorID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblBranchesData", "Beauty", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblOptions", "SpecialVersion", adBoolean, adColNullable, , , " ", False, True

    Cn.Execute "update transactiontypes set TransactionTypeName='.' ,TransactionEnglishName='.'  where transaction_type=44"

    Cn.Execute " update transactiontypes  set TransactionTypeName='√Ê«„— «·»Ì⁄  «·„»Ì⁄« ' ,TransactionEnglishName='Sales Order'  where transaction_type=6"

    If DB_CreateTable("TblFarmColors", True, "ID", False) = True Then
        DB_CreateField "TblFarmColors", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblFarmColors", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
    End If

    DB_CreateField "TblFarmAnimalRegister", "ColorID", adInteger, adColNullable, , , "    ", False, True

    If DB_CreateTable("TblLensesTypes", True, "ID", False) = True Then
        DB_CreateField "TblLensesTypes", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblLensesTypes", "NameE", adVarWChar, adColNullable, 400, , "      ", False, True, , True
        DB_CreateField "TblLensesTypes", "GroupID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblLensesTypes", "UnitID", adInteger, adColNullable, , , "    ", False, True
        DB_CreateField "TblLensesTypes", "Flag", adInteger, adColNullable, , , "    ", False, True
           
        DB_CreateField "TblLensesTypes", "FromSPH", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblLensesTypes", "TOSPH", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblLensesTypes", "FROMCYL", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblLensesTypes", "TOCYL", adDouble, adColNullable, , , "    ", False, True
           
    End If

    DB_CreateField "Transaction_Details", "RCL", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Transaction_Details", "LCL", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "TblLensesTypes", "Price", adDouble, adColNullable, , , "    ", False, True
 
    DB_CreateField "Transaction_Details", "ItemBalance2", adDouble, adColNullable, , , "    ", False, True

    DB_CreateField "Transactions", "IsTransfere", adBoolean, adColNullable, , , "    ", False, True

    DB_CreateField "Transaction_Details", "RequestTypeNo", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "Transaction_Details", "StoreIDAvi", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "Transaction_Details", "LCL", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblFarmAnimalRegister", "DoctorID", adInteger, adColNullable, , , "    ", False, True

    DB_CreateField "TblLensesTypes", "BrandsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "TypeItemsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DesignID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "CollectionsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ShapesID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ShapesNewID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "MaterialID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "SexID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "AGEID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "CylinderID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "SphereID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "PackingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "UsageID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "GroupEyeID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "BaseCurveID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "NationalityID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ServiceID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "BreakingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "LightAdaptationID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DIAMID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "IndexsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "CoatingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DivisionID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "OriginID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "ColorID11", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "MasterType", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblOptions", "IsShowLensesDetails", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "BrandsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "TypeItemsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DesignID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "CollectionsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ShapesID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ShapesNewID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "MaterialID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "SexID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "AGEID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "CylinderID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "SphereID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "PackingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "UsageID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "GroupEyeID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "BaseCurveID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "NationalityID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "ServiceID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "BreakingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "LightAdaptationID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DIAMID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "IndexsID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "CoatingID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "DivisionID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "TblLensesTypes", "OriginID", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "ColorID11", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblLensesTypes", "MasterType", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblStore", "IsLab", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "TransferMoveID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "PurchaseRequestID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "IsOrderOut", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "StoreIDAvi2", adInteger, adColNullable, , , "", False, True

    DB_CreateField "TblStore", "IsLab", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "TransferMoveID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "PurchaseRequestID", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "IsOrderOut", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "IsFinish", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "StoreIDAvi2", adInteger, adColNullable, , , "", False, True
    DB_CreateField "Groups", "IsTransfere", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Groups", "IsTransfere", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblBranchesData", "StoreId", adInteger, adColNullable, , , " ???    ", False, True

    DB_CreateField "Transaction_Details", "IsTransfere", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "IsTransfere", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "IsTransfere", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "Transaction_Details", "ReadSPH", adVarWChar, adColNullable, 255, , "", False, True
    DB_CreateField "Transaction_Details", "ReadCYL", adVarWChar, adColNullable, 255, , "", False, True
    DB_CreateField "Transaction_Details", "LReadSPH", adVarWChar, adColNullable, 255, , "", False, True
    DB_CreateField "Transaction_Details", "LReadCYL", adVarWChar, adColNullable, 255, , "", False, True

    DB_CreateField "TblBranchesData", "StoreId", adInteger, adColNullable, , , " ???    ", False, True

    If DB_CreateTable("TblFixedAssestTmpValue", True, "ID", True) = True Then
        DB_CreateField "TblFixedAssestTmpValue", "FixedId", adInteger, adColNullable, , , , False, True
        DB_CreateField "TblFixedAssestTmpValue", "TransID", adInteger, adColNullable, , , , False, True
        DB_CreateField "TblFixedAssestTmpValue", "NetValue", adDouble, adColNullable, , , , False, True
    End If

    DB_CreateField "Transaction_Details", "ProfitType", adInteger, adColNullable, , , , False, True
    DB_CreateField "Transaction_Details", "ProfitValue", adDouble, adColNullable, , , , False, True

    DB_CreateField "branches", "a167", adVarWChar, adColNullable, 255, , "", False, True
    DB_CreateField "branches", "a168", adVarWChar, adColNullable, 255, , "", False, True

    DB_CreateField "branches", "a211", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "branches", "a212", adVarWChar, adColNullable, 250, , "", False, True, , True

    DB_CreateField "Transaction_Details", "ProfitType", adInteger, adColNullable, , , , False, True
    DB_CreateField "Transaction_Details", "ProfitValue", adDouble, adColNullable, , , , False, True
    DB_CreateField "Transaction_Details", "NetProfit", adDouble, adColNullable, , , , False, True

    DB_CreateField "branches", "a167", adVarWChar, adColNullable, 255, , "", False, True

    add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "50,'«· ’—›«  «·⁄ﬁ«—Ì…  ','Pay Contract'", "ID", 50

    DB_CreateField "notes_all", "AkarPayCheck", adInteger, adColNullable, , , , False, True

    DB_CreateField "tblItems", "LensesTypesID", adInteger, adColNullable, , , , False, True

    sql = "    DROP FUNCTION QryItemsTransactionsTotals2" & CHR(13)
    Cn.Execute sql

    s = " Create FUNCTION QryItemsTransactionsTotals2"
    s = s & " ("
    s = s & "     @TransType          INT = 0,"
    s = s & "     @TransType2         INT = 0,"
    s = s & "     @TransType3         INT = 0,"
    s = s & "     @FromDate           DATETIME,"
    s = s & "     @ToDate             DATETIME,"
    s = s & "     @ItemID             AS integer,"
    s = s & "     @Transaction_ID     AS FLOAT = NULL"
    s = s & " )"
    s = s & " RETURNS @xTable TABLE"
    s = s & "         ("
    s = s & "             ItemID INT,"
    s = s & "             ItemCode NVARCHAR(50),"
    s = s & "             ItemName NVARCHAR(4000),"
    s = s & "             GroupID INT,"
    s = s & "             Total FLOAT,"
    s = s & "             totalqty Float"
    s = s & "         )"
    s = s & " AS"

    s = s & " Begin"
    s = s & "     INSERT @xTable"
    s = s & "     SELECT ItemID,"
    s = s & "            ItemCode,"
    s = s & "            ItemName,"
    s = s & "            GroupID,"
    s = s & "            SUM(Total)     AS Totals,"
    s = s & "            SUM(Quantity)  As totalqty"
    s = s & "     FROM   ("
    s = s & "                SELECT TblItems.ItemID,"
    s = s & "                       TblItems.ItemCode,"
    s = s & "                       TblItems.ItemName,"
    s = s & "                       TblItems.GroupID,"
    s = s & "                       'Total' = CASE"
    s = s & "                                      WHEN ItemDiscountType = 1"
    s = s & "                OR ItemDiscountType = 0 THEN Transaction_Details.Quantity * Transaction_Details.Price"
    s = s & "                   WHEN ItemDiscountType = 2 THEN ("
    s = s & "                       (Transaction_Details.Quantity * Transaction_Details.Price) -ItemDiscount"
    s = s & "                   )"
    s = s & "                   WHEN ItemDiscountType = 3 THEN (Transaction_Details.Quantity * Transaction_Details.Price) * (1 -(ItemDiscount / 100))"
    s = s & "                   ELSE 0"
    s = s & "                   END,"
    s = s & "                Transaction_Details.Quantity"
    s = s & "                FROM dbo.TblItems INNER JOIN dbo.Transaction_Details ON dbo.TblItems.ItemID ="
    s = s & "                dbo.Transaction_Details.Item_ID INNER JOIN"
    s = s & "                dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "                WHERE ("
    s = s & "                    Transactions.Transaction_Type = @TransType"
    s = s & "                    OR Transactions.Transaction_Type = @TransType2"
    s = s & "                    OR Transactions.Transaction_Type = @TransType3"
    s = s & "                    OR Transactions.Transaction_Type = 34"
'    s = s & "                    OR Transactions.Transaction_Type = 11"
    s = s & "                    OR Transactions.Transaction_Type = 15"
    s = s & "                )"
    s = s & "                AND Transactions.Transaction_Date >= @FromDate"
    s = s & "                AND Transactions.Transaction_Date <= @TODate"
          
    s = s & "                AND Transactions.Transaction_ID <> ISNULL(@Transaction_ID, Transactions.Transaction_ID)"
    s = s & "            )                 DrivTable"
    s = s & "     Group By"
    s = s & "            ItemID,"
    s = s & "            ItemCode,"
    s = s & "            ItemName,"
    s = s & "            GroupID"
    
    s = s & "     Return"
    s = s & " End"
    
    db_createOrUpdateFuctionSQL "QryItemsTransactionsTotals2", s
    DB_CreateField "Accounts", "account_serial1", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "account_name1", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "Account_NameEng1", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "Accounts", "account_serial2", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "account_name2", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "Account_NameEng2", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "Accounts", "account_serial3", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "account_name3", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "Accounts", "Account_NameEng3", adVarWChar, adColNullable, 255, , "      ", False, True, , True

    DB_CreateField "Transactions", "FirstEntryDateDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    
    'Cn.Execute " TRIGGER [dbo].[trgFirstEntryDate]"
'    s = "    CREATE TRIGGER dbo.trgFirstEntryDate ON dbo.Transactions"
'    s = s & "  AFTER insert"
'    s = s & "  AS"
'    s = s & "   update dbo.Transactions"
'    s = s & "   Set FirstEntryDateDate = GetDate()"
'    s = s & "   FROM Inserted i"
'    s = s & "  Where dbo.Transactions.Transaction_ID = i.Transaction_ID"
'    Cn.Execute s
    
    

s = "IF NOT EXISTS (SELECT * FROM sys.triggers WHERE name = 'trgFirstEntryDate') "
s = s & "BEGIN "
s = s & "CREATE TRIGGER dbo.trgFirstEntryDate ON dbo.Transactions "
s = s & "AFTER INSERT "
s = s & "AS "
s = s & "UPDATE dbo.Transactions "
s = s & "SET FirstEntryDateDate = GETDATE() "
s = s & "FROM Inserted i "
s = s & "WHERE dbo.Transactions.Transaction_ID = i.Transaction_ID "
s = s & "END"

Cn.Execute s
On Error Resume Next
Cn.Execute "DROP TRIGGER dbo.trgFirstEntryDate"
On Error GoTo 0

Cn.Execute _
"EXEC('CREATE TRIGGER dbo.trgFirstEntryDate ON dbo.Transactions " & _
"AFTER INSERT AS BEGIN SET NOCOUNT ON; " & _
"UPDATE t SET FirstEntryDateDate = ISNULL(t.FirstEntryDateDate, GETDATE()) " & _
"FROM dbo.Transactions t INNER JOIN Inserted i ON t.Transaction_ID = i.Transaction_ID; END')"



    DB_CreateField "TblVATAvowal", "HideakarExitVat1", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblVATAvowal", "HideakarExitVat2", adBoolean, adColNullable, , , "                ", False, True

    DB_CreateField "TblCustomersLocations", "TXTDOBLOcation", adVarWChar, adColNullable, 255, , "      ", False, True, , True
    DB_CreateField "TblCustomersLocations", "CUsDOB", adDBTimeStamp, adColNullable, , , "      ", False, True

    DB_CreateField "project_billl", "PostedDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "project_billl", "Approved", adBoolean, adColNullable, , , "???? ?? ??", False, True
      
    DB_CreateField "project_billl", "Posted", adInteger, adColNullable, , , "      ", False, True
      
    If DB_CreateTable("TblCaptinTrans", True, "id ", False) = True Then
        DB_CreateField "TblCaptinTrans", "RecordDate", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
        DB_CreateField "TblCaptinTrans", "NoteSerial", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblCaptinTrans", "NoteSerial1", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True

        DB_CreateField "TblCaptinTrans", "IsVat", adBoolean, adColNullable, , , "      ", False, True
        DB_CreateField "TblCaptinTrans", "NoteID", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblCaptinTrans", "BankID", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblCaptinTrans", "BranchID", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblCaptinTrans", "UserID", adInteger, adColNullable, , , "      ", False, True
       
    End If



    'add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 23001 ,'ﬁÌœ Õ—ﬂ«  «· ‰ﬁ·' ,'      ' ", "NotesType", 23001

    sql = "DROP FUNCTION GetTotEx"
    Cn.Execute sql
    sql = " CREATE FUNCTION GetTotEx(@ID integer ,@IDH integer )"
    sql = sql & "  RETURNS Float"
    sql = sql & " AS"
    sql = sql & " Begin"
    sql = sql & " RETURN ( SELECT     SUM( TOTEX*(1+FATYou/100)  ) AS SumTotEx"
    sql = sql & " FROM         dbo.project_bill_details INNER JOIN"
    sql = sql & "                  dbo.project_billl ON dbo.project_bill_details.bill_id = dbo.project_billl.id"
    sql = sql & "  Where (dbo.project_bill_details.oprid = @ID) And (dbo.project_bill_details.oprid <> 0)"
    sql = sql & "        AND (dbo.project_billl.id <  @IDH)"
    sql = sql & "   )"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "GetTotEx", sql
    
    sql = "  DROP FUNCTION GetQuntExc"
    Cn.Execute sql
    sql = " CREATE FUNCTION GetQuntExc(@ID integer ,@IDH integer)"
    sql = sql & "  RETURNS Float"
    sql = sql & " AS"
    sql = sql & " Begin"
    sql = sql & " RETURN ( SELECT     SUM(dbo.project_bill_details.quntExc) AS SumQuntExc"
    sql = sql & " FROM         dbo.project_bill_details INNER JOIN"
    sql = sql & "                  dbo.project_billl ON dbo.project_bill_details.bill_id = dbo.project_billl.id"
    sql = sql & "  Where (dbo.project_bill_details.oprid = @ID) And (dbo.project_bill_details.oprid <> 0)"
    sql = sql & "        AND (dbo.project_billl.id <  @IDH)"
    sql = sql & "   )"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "GetQuntExc", sql
    
    sql = "  DROP FUNCTION GetOLDPerforValue " & CHR(13)
    Cn.Execute sql
    sql = " CREATE FUNCTION GetOLDPerforValue(@ProjectNo integer ,@id integer)"
    sql = sql & "  RETURNS Float"
    sql = sql & " AS"
    sql = sql & " Begin"
    sql = sql & " RETURN ( select  sum(PerforValue)  as PerforValueOLdTotal  "
    sql = sql & " FROM         dbo.project_billl "
     
    sql = sql & "  Where (dbo.project_billl.project_no = @ProjectNo) "
    sql = sql & "        AND (dbo.project_billl.id <  @ID)"
    sql = sql & "   )"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "GetOLDPerforValue", sql
    
    sql = "  DROP FUNCTION GetOLDPerforValuebysubContractorId " & CHR(13)
    Cn.Execute sql
    sql = " CREATE FUNCTION GetOLDPerforValuebysubContractorId(@ProjectNo integer ,@id integer,@subContractorId AS integer)"
    sql = sql & "  RETURNS Float"
    sql = sql & " AS"
    sql = sql & " Begin"
    sql = sql & " RETURN ( select  sum(PerforValue)  as GetOLDPerforValuebysubContractorId  "
    sql = sql & " FROM         dbo.project_billl "
     
    sql = sql & "  Where (dbo.project_billl.project_no = @ProjectNo) "
    sql = sql & "        AND (dbo.project_billl.id <  @ID)"
    sql = sql & "        AND (dbo.project_billl.subContractorId =  @subContractorId)"
    sql = sql & "   )"
    sql = sql & " End"
    db_createOrUpdateFuctionSQL "GetOLDPerforValuebysubContractorId", sql
    
    DB_CreateField "TblApprovalDef", "DepartmentID", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "notes_all", "DepartmentID", adInteger, adColNullable, , , "      ", False, True
    
    DB_CreateField "tblContractInsAllocationsDetails", "nextinstalldateH", adVarWChar, adColNullable, 10, , "", False, True, , True
    DB_CreateField "tblContractInsAllocationsDetails", "nextinstalldate", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
    DB_CreateField "tbloptions", "DefaultQtyTrans", adVarWChar, adColNullable, 5, , "", False, True, , True
    DB_CreateField "TblContract", "UserID", adInteger, adColNullable, , , , False, True

    DB_CreateField "TblUsers", "CanAcreditRsContract", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "CanIsShamel", adBoolean, adColNullable, , , "", False, True
    
    
    DB_CreateField "TblUsers", "OPenShortInvoice", adBoolean, adColNullable, , , "", False, True
    
    
   
 
        
    DB_CreateField "project_bill_details", "LineDiscountPercent", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "LineDiscount", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "linenetaftermainDiscount", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "linenetaftermainDiscountBeforevat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "LineVat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "linenetaftermainDiscountWithvat", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "PerforVLineDiscount", adDouble, adColNullable, , , "", False, True
    DB_CreateField "project_bill_details", "LineFinal", adDouble, adColNullable, , , "", False, True
                  

                   
    'DB_CreateField "TblUsers", "OPenShortInvoice", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "MonyeIssueVchrNoMust", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "POMustentryAndBillMustEntry", adBoolean, adColNullable, , , "", False, True

    DB_CreateField "TblExchange", "salary_or_advance", adInteger, adColNullable, , , , False, True
    
    DB_CreateField "groups", "ActivityTypeId", adDouble, adColNullable, , , "", False, True
     
    DB_CreateField "tbloptions", "WorkWithLINKEDiActivity", adBoolean, adColNullable, , , "", False, True
     
    ageingFunc
         
    UpdateCostriceProcedure
        
 
    DB_CreateField "TransactionTypes", "projectInclude", adBoolean, adColNullable, , , "", False, True
        
    DB_CreateField "Transactions", "OverProject", adDouble, adColNullable, , , "", False, True
    DB_CreateField "Transaction_Details", "OverProject", adDouble, adColNullable, , , "", False, True
    DB_CreateField "ApprovalData", "OverProject", adDouble, adColNullable, , , "", False, True
    
    DB_CreateField "TblContractInstallments", "VATYou1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATYou2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue2", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "VATYou1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATYou2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue2", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "CountDay1", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CountDay2", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "VATYou1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATYou2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CostDay", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "CountDay1", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CountDay2", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CountDaysTotal", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "IsChangVat", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "VATYou1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATYou2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue1", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "VATValue2", adDouble, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CostDay", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "CountDay1", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CountDay2", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "CountDaysTotal", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "IsChangVat", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "TblContractInstallments", "NoteIdDiff", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblContractInstallments", "NoteSerialDiff", adInteger, adColNullable, , , "      ", False, True

    DB_CreateField "Transaction_Details", "thickness", adDouble, adColNullable, , , "      ", False, True

    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 29801 ,'ﬁÌœ Õ—ﬂ«  Õ—ﬂ«  ›—Êﬁ«  «·÷—Ì»…' ,'      ' ", "NotesType", 29801

    'wawa
    DB_CreateField "tblGeneralCashingDetails", "salesValue", adDouble, adColNullable, , , "      ", False, True

    DB_CreateField "TblUsers", "HideTbarInPos", adBoolean, adColNullable, , , "", False, True
   
   

End Function

Function ageingFunc()
On Error Resume Next


DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Notes", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True


DB_CreateField "TblBanksDeposite", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblBanksDepositeDetails", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblBanksCollect", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblBanksCollectdetails", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "TblCaptinTrans", "DueDate", adDBTimeStamp, adColNullable, , , "", False, True



    If DB_CreateTable("TblAging", True, "id ", True) = True Then
        DB_CreateField "TblAging", "RecordDate", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
        DB_CreateField "TblAging", "DueDate", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
        DB_CreateField "TblAging", "DueDate2", adDBTimeStamp, adColNullable, , , " «—ÌŒ  «·⁄„·Ì…  ", False, True
        DB_CreateField "TblAging", "Account_Code", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblAging", "NoteSerial1", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblAging", "CusName", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblAging", "TransactionTypeName", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
         DB_CreateField "TblAging", "DiffDate", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
         DB_CreateField "TblAging", "TransactionTypeName2", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        
        DB_CreateField "TblAging", "Balance", adDouble, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "TransNet", adDouble, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "PayedValue", adDouble, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "StillAmount", adDouble, adColNullable, , , "      ", False, True
        
        DB_CreateField "TblAging", "Transaction_Type", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "AGEID", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "CusID", adInteger, adColNullable, , , "      ", False, True
        DB_CreateField "TblAging", "Credit_Or_Debit", adInteger, adColNullable, , , "      ", False, True
        

           DB_CreateField "TblAging", "BranchID", adInteger, adColNullable, , , "      ", False, True
       DB_CreateField "TblAging", "UserID", adInteger, adColNullable, , , "      ", False, True

    End If
  
 DB_CreateField "TblAging", "DiffDate", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblAging", "NoteSerial", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblAging", "[NAME]", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblAging", "[From]", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblAging", "[To]", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblAging", "[Color]", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True


 
   Dim sql As String
    sql = sql & "    DROP VIEW RptLedger_Sub2" & CHR(13)

    Cn.Execute sql

 sql = ""
 sql = sql & "   Create VIEW RptLedger_Sub2 AS " & CHR(13)
 sql = sql & "   SELECT dbo.Notes.ChqueNum," & CHR(13)
 sql = sql & "          dbo.Notes.ManualNo," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE," & CHR(13)
 sql = sql & "          dbo.ACCOUNTS.Account_Name," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No," & CHR(13)
 sql = sql & "          dbo.TblNotesTypes.NotesTypeName," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.UserID," & CHR(13)
 sql = sql & "          dbo.TblUsers.UserName," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.OperaID," & CHR(13)
 sql = sql & "          dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID," & CHR(13)
sql = sql & "                 dbo.Transactions.Transaction_Serial," & CHR(13)
sql = sql & "                 dbo.Transactions.Transaction_Date," & CHR(13)
sql = sql & "                 dbo.TransactionTypes.TransactionTypeName," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID," & CHR(13)
sql = sql & "                 dbo.Notes.NoteDate," & CHR(13)
sql = sql & "                 dbo.Notes.NoteType," & CHR(13)
sql = sql & "                 dbo.Notes.NoteSerial," & CHR(13)
sql = sql & "                 DOUBLE_ENTREY_VOUCHERS.value  Note_Value," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Account_Serial," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Account_NameEng," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Parent_Account_Code," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.opening_balance," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.opening_balance_type," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Branch," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Sum_account," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.cost_center," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.currenct_code," & CHR(13)
sql = sql & "                 dbo.Notes.Remark," & CHR(13)
       
sql = sql & "                 dbo.Notes.note_value_by_characters," & CHR(13)
sql = sql & "                 dbo.Notes.foxy_no," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1," & CHR(13)
sql = sql & "                 dbo.TblNotesTypes.NotesTypeNamee," & CHR(13)
sql = sql & "                 dbo.TransactionTypes.TransactionEnglishName," & CHR(13)
sql = sql & "                 dbo.Notes.NoteSerial1," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.branch_id," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.ActivityTypeId," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.notes_all," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.branch_name," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.branch_namee," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.Posted," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.currency," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.rate," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.RegionID," & CHR(13)
sql = sql & "                 dbo.TblSection.name," & CHR(13)
sql = sql & "                 dbo.TblSection.namee," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.project_id," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.projectid," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.operid," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.pandid," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid," & CHR(13)
sql = sql & "                 dbo.TblAqar.aqarname," & CHR(13)
sql = sql & "                 dbo.TblAqar.aqarNo," & CHR(13)
sql = sql & "                 DueDate = IsNull(DOUBLE_ENTREY_VOUCHERS.DueDate,IsNull(Notes.DueDate,RecordDate))"
sql = sql & "          From dbo.TblAqar" & CHR(13)
sql = sql & "                 RIGHT OUTER JOIN dbo.TblBranchesData" & CHR(13)
sql = sql & "                 INNER JOIN dbo.TblUsers" & CHR(13)
sql = sql & "                 INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS" & CHR(13)
sql = sql & "                      ON  dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID" & CHR(13)
sql = sql & "                      ON  dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id" & CHR(13)
sql = sql & "                      ON  dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.ACCOUNTS" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TblSection" & CHR(13)
sql = sql & "                      ON  dbo.TblBranchesData.RegionID = dbo.TblSection.Id" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.Notes" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TblNotesTypes" & CHR(13)
sql = sql & "                      ON  dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.Transactions" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TransactionTypes" & CHR(13)
sql = sql & "                      ON  dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type" & CHR(13)
sql = sql & "          Where (dbo.DOUBLE_ENTREY_VOUCHERS.Posted Is Null)" & CHR(13)


sql = sql & "          Union all" & CHR(13)
sql = sql & "          SELECT dbo.Notes.ChqueNum," & CHR(13)
sql = sql & "                 dbo.Notes.ManualNo," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_ID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Credit_Or_Debit," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.[Value] AS DEV_Value," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDateH," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Description AS DEV_DES," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Double_Entry_Vouchers_Descriptione AS DevDESE," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Account_Name," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No," & CHR(13)
sql = sql & "                 NotesTypeName = 'ﬁÌœ «›  «ÕÌ'," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.UserID," & CHR(13)
sql = sql & "                 dbo.TblUsers.UserName," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDate," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.ReceiptID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.OperaID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Transaction_ID," & CHR(13)
sql = sql & "                 dbo.Transactions.Transaction_Serial," & CHR(13)
sql = sql & "                 dbo.Transactions.Transaction_Date," & CHR(13)
sql = sql & "                 dbo.TransactionTypes.TransactionTypeName," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.PostedDate," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.PostedUserID," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Interval_ID," & CHR(13)
sql = sql & "                 dbo.Notes.NoteDate," & CHR(13)
sql = sql & "                 dbo.Notes.NoteType," & CHR(13)
sql = sql & "                 dbo.Notes.NoteSerial," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.value Note_Value," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Account_Serial," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Account_NameEng," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Parent_Account_Code," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.opening_balance," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.opening_balance_type," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Branch," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.Sum_account," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.cost_center," & CHR(13)
sql = sql & "                 dbo.ACCOUNTS.currenct_code," & CHR(13)
sql = sql & "                 dbo.Notes.Remark," & CHR(13)
       
sql = sql & "                 dbo.Notes.note_value_by_characters," & CHR(13)
sql = sql & "                 dbo.Notes.foxy_no," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.DEV_ID_Line_No1," & CHR(13)
sql = sql & "                 dbo.TblNotesTypes.NotesTypeNamee," & CHR(13)
sql = sql & "                 dbo.TransactionTypes.TransactionEnglishName," & CHR(13)
sql = sql & "                 dbo.Notes.NoteSerial1," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.ActivityTypeId," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.notes_all," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.branch_name," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.branch_namee," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.Posted," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.valuee AS DEV_ValueE," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.currency," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.rate," & CHR(13)
sql = sql & "                 dbo.TblBranchesData.RegionID," & CHR(13)
sql = sql & "                 dbo.TblSection.name," & CHR(13)
sql = sql & "                 dbo.TblSection.namee," & CHR(13)
sql = sql & "                 '' as DescAccount," & CHR(13)
sql = sql & "                  '' as NextAccount_Code," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.project_id," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.opr_fullcode," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.projectid," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.operid," & CHR(13)
sql = sql & "                 dbo.DOUBLE_ENTREY_VOUCHERS1.pandid," & CHR(13)
sql = sql & "                 0 as Aqarid," & CHR(13)
sql = sql & "                 '' as aqarname," & CHR(13)
sql = sql & "                 0 as aqarNo," & CHR(13)

sql = sql & "                 DueDate = IsNull(DOUBLE_ENTREY_VOUCHERS1.DueDate,IsNull(Notes.DueDate,RecordDate)) " & CHR(13)
sql = sql & "          From " & CHR(13)
sql = sql & "                 dbo.TblBranchesData" & CHR(13)
sql = sql & "                 INNER JOIN dbo.TblUsers" & CHR(13)
sql = sql & "                 INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS1" & CHR(13)
sql = sql & "                      ON  dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS1.UserID" & CHR(13)
sql = sql & "                      ON  dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id" & CHR(13)

sql = sql & "                 LEFT OUTER JOIN dbo.ACCOUNTS" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code = dbo.ACCOUNTS.Account_Code" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TblSection" & CHR(13)
sql = sql & "                      ON  dbo.TblBranchesData.RegionID = dbo.TblSection.Id" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.Notes" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TblNotesTypes" & CHR(13)
sql = sql & "                      ON  dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS1.Notes_ID = dbo.Notes.NoteID" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.Transactions" & CHR(13)
sql = sql & "                      ON  dbo.DOUBLE_ENTREY_VOUCHERS1.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
sql = sql & "                 LEFT OUTER JOIN dbo.TransactionTypes" & CHR(13)
sql = sql & "                      ON  dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type" & CHR(13)
sql = sql & "          Where (dbo.DOUBLE_ENTREY_VOUCHERS1.Posted Is Null)" & CHR(13)
 
    Cn.Execute sql
End Function
Public Function mostafa()
  

End Function

Public Function updateFuncSqaccountMovesl()

 On Error Resume Next
 
    Dim sql As String
DB_CreateField "TblContractInstallments", "VATYou1", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "VATYou2", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "VATValue1", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "VATValue2", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "CostDay", adDouble, adColNullable, , , "      ", False, True

DB_CreateField "TblContractInstallments", "CountDay1", adInteger, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "CountDay2", adInteger, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "CountDaysTotal", adInteger, adColNullable, , , "      ", False, True

DB_CreateField "TblContractInstallments", "IsChangVat", adInteger, adColNullable, , , "      ", False, True



DB_CreateField "TblContractInstallments", "NoteIdDiff", adInteger, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "NoteSerialDiff", adInteger, adColNullable, , , "      ", False, True


add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 29801 ,'ﬁÌœ Õ—ﬂ«  ›—Êﬁ«  «·÷—Ì»…' ,'      ' ", "NotesType", 29801




DB_CreateField "TblContract", "FATYou22", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblContractInstallments", "DiffAmount", adDouble, adColNullable, , , "      ", False, True

DB_CreateField "TblUsers", "ReportName2", adVarWChar, adColNullable, 255, , "", False, True, , True

           
   DB_CreateField "tblaqar", "FromPlanneddate", adDBTimeStamp, adColNullable, , , "      ", False, True
  DB_CreateField "tblaqar", "FromPlanneddateH", adVarWChar, adColNullable, 10, , "      ", False, True, , True
 DB_CreateField "tblaqar", "ToPlanneddate", adDBTimeStamp, adColNullable, , , "      ", False, True
  DB_CreateField "tblaqar", "ToPlanneddateH", adVarWChar, adColNullable, 10, , "      ", False, True, , True
  
  
  DB_CreateField "tblaqar", "PlotNo", adVarWChar, adColNullable, 255, , "      ", False, True, , True
  DB_CreateField "tblaqar", "Planned", adVarWChar, adColNullable, 255, , "      ", False, True, , True
  DB_CreateField "tblaqar", "PlotNo", adVarWChar, adColNullable, 255, , "      ", False, True, , True
  DB_CreateField "tblaqar", "DisountAmount", adDouble, adColNullable, , , "    ", False, True
    

                

DB_CreateField "TblPrintBarCode", "Suppliername", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblPrintBarCode", "Suppliercode", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblPrintBarCode", "Zcode", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "TblPrintBarCode", "Zcode128", adVarWChar, adColNullable, 255, , "      ", False, True, , True


DB_CreateField "TblContractInstallments", "VATValue1Com", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "TblContractInstallments", "VATValue2Com", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "TblContractInstallments", "Commissions2", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblPrintBarCode", "qty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "SubcontractorContract", "Prefix", adInteger, adColNullable, , , "  ", False, True
DB_CreateField "SubcontractorContract", "NoteSerial1", adInteger, adColNullable, , , "  ", False, True
    If DB_CreateTable("SubcontractorContract", True, "ID", True) = True Then
            DB_CreateField "SubcontractorContract", "Commissions2", adDouble, adColNullable, , , "    ", False, True
            
            DB_CreateField "SubcontractorContract", "FATYou", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "FATValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "TotalValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "Period", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "AccountCodeVat", adVarWChar, adColNullable, 55, , "C?C??   ", False, True, , True
            
            
            DB_CreateField "SubcontractorContract", "project_name", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "NoteSerial", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "project_no", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "End_user_name", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "End_user_account", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "bill_to", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "bill_type", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "BillNo", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "revenue_account", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "bill_type", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "SubcontractorContract", "note_id", adInteger, adColNullable, , , "  ", False, True
            DB_CreateField "SubcontractorContract", "Branch_NO", adInteger, adColNullable, , , "  ", False, True
            DB_CreateField "SubcontractorContract", "UserID", adInteger, adColNullable, , , "  ", False, True
            
            
            DB_CreateField "SubcontractorContract", "subContractorId", adInteger, adColNullable, , , "  ", False, True
            DB_CreateField "SubcontractorContract", "PeriodType", adInteger, adColNullable, , , "  ", False, True
            
            DB_CreateField "SubcontractorContract", "total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "Results", adDouble, adColNullable, , , "    ", False, True
            
            DB_CreateField "SubcontractorContract", "bill_date", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "SubcontractorContract", "dueDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            
            DB_CreateField "SubcontractorContract", "ExPercen", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "ExPercenID", adInteger, adColNullable, , , "  ", False, True
            DB_CreateField "project_bill_details", "ExPercen", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreVAT", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaVAT", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaTotal", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaPayed", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaRemain", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaTransPyed", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaNet", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PreBalaVATYu", adDouble, adColNullable, , , "    ", False, True
            
            
            DB_CreateField "SubcontractorContract", "StartDateProje", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "SubcontractorContract", "LineVat", adDBTimeStamp, adColNullable, , , "      ", False, True
            
            DB_CreateField "SubcontractorContract", "SumVATLine", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "SumValueLine", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "NetValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PerforValue", adDouble, adColNullable, , , "    ", False, True
            
            DB_CreateField "SubcontractorContract", "NetValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "SubcontractorContract", "PerforValue", adDouble, adColNullable, , , "    ", False, True
            
            
    End If




'Projects.show


        If DB_CreateTable("SubcontractorContract2", True, "ID", True) = True Then
        
DB_CreateField "SubcontractorContract2", "projectName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "SubcontractorContract2", "FullCode", adVarWChar, adColNullable, 4000, , "      ", False, True, , True

                DB_CreateField "SubcontractorContract2", "project_no", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
                DB_CreateField "SubcontractorContract2", "item", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
                DB_CreateField "SubcontractorContract2", "item_id", adVarWChar, adColNullable, 50, , "      ", False, True, , True
                DB_CreateField "SubcontractorContract2", "item_unit", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
                DB_CreateField "SubcontractorContract2", "Period", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "cost", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "exe", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "percentage", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Price", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Pre_Quantity", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Quantity", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "bill_id", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Unit_id", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "line_no", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "LineVat2", adDBTimeStamp, adColNullable, , , "      ", False, True
                DB_CreateField "SubcontractorContract2", "project_id", adInteger, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "Curr_Quantity", adInteger, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "Pre_Value", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Pre_Percent", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Curr_value", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "curr_Percent", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "tot_quantity", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "tot_value", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "tot_percent", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "qty", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "total", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "discount", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "net", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "quntExc", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "totEx", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "discountEXE", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "NetExe", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "percentage1", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "Pre_Percent1", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "tot_percent1", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "QtyApprov", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "TotalApprov", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "PriceApprov", adDouble, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "DiscApprov", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "NetApprov", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "ExPercen", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "LineDiscountPercent", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "LineDiscount", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "linenetaftermainDiscount", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "linenetaftermainDiscountBeforevat", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "linenetaftermainDiscountWithvat", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "PerforVLineDiscount", adDouble, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "LineFinal", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "qtySubContractor", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "costSubContractor", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "OLDTotalwithVat", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "CurrenttotalWithvat", adDouble, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "Totalwitvat", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "CurrenttotalWithvat", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "oldPerforValue", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "totalPerforValue", adDouble, adColNullable, , , "    ", False, True
                
                DB_CreateField "SubcontractorContract2", "oprid", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "PrMainDesID", adInteger, adColNullable, , , "    ", False, True
                DB_CreateField "SubcontractorContract2", "exedate", adDBTimeStamp, adColNullable, , , "      ", False, True
        End If

 

    
   DB_CreateField "tbloptions", "amlaketbatrentOnly", adBoolean, adColNullable, , , "", False, True

  add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "51,'⁄„Ê·«  «œ«—… «·⁄ﬁ«—    ','Pay Contract'", "ID", 51
   
DB_CreateField "TblGroupItemProductLineUsersset", "ProgramUsername", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "logfile", "ProgramUsername", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "logfile", "Computername", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True
DB_CreateField "logfile", "ComputerUsername", adVarWChar, adColNullable, 255, , "C?C??   ", False, True, , True

DB_CreateField "transactions", "chkDone", adInteger, adColNullable, , , "    ", False, True



If DB_CreateTable("tblBoxesTemp", True, "ID", True) = True Then
            DB_CreateField "tblBoxesTemp", "Serial", adDouble, adColNullable, , , "    ", False, True
            
            DB_CreateField "tblBoxesTemp", "BoxID", adInteger, adColNullable, , , "    ", False, True
            DB_CreateField "tblBoxesTemp", "BoxName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            
            DB_CreateField "tblBoxesTemp", "DebitValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblBoxesTemp", "CreditValue", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblBoxesTemp", "Period", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "tblBoxesTemp", "Type", adVarWChar, adColNullable, 55, , "C?C??   ", False, True, , True
            
            
            
            
End If
DB_CreateField "tblaqar", "NOOFYears", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblaqar", "TypeDate", adInteger, adColNullable, , , "    ", False, True


DB_CreateField "tblaqar", "NoteID", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "tblaqar", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True
DB_CreateField "tblaqar", "NoteSerial1", adVarWChar, adColNullable, 255, , "      ", False, True, , True

add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 29802 ,'ﬁÌœ ⁄ﬁÊœ ·„·«ﬂ    ' ,'      ' ", "NotesType", 29802
 DB_CreateField "TblOptions", "NotAllowStockNegativeInternal", adBoolean, adColNullable, , , "                ", False, True


DB_CreateField "tblaqar", "TxtContValueWithout", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblaqar", "TxtFATValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblaqar", "TxtFATYou", adDouble, adColNullable, , , "    ", False, True
 
 
 
DB_CreateField "TblAqrOwin", "valuewithout", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAqrOwin", "VatPerc", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAqrOwin", "VatValue", adDouble, adColNullable, , , "    ", False, True



DB_CreateField "TblOptions", "MustEnterNewNo", adBoolean, adColNullable, , , "                ", False, True
 add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "52,'    ⁄ﬁÊœ «·„·«ﬂ     ','Pay Contract'", "ID", 52
   
   DB_CreateField "tblaqar", "ComResid", adDouble, adColNullable, , , "    ", False, True
   
   
  DB_CreateField "transactions", "VstReverse", adBoolean, adColNullable, , , "", False, True
  DB_CreateField "Transaction_Details", "VstReverse", adBoolean, adColNullable, , , "", False, True
  
  
  
  DB_CreateField "TblVATAvowal", "TxtBillVstReverse", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "TblVATAvowal", "TxtBillVstReverseREt", adDouble, adColNullable, , , "    ", False, True
  
  DB_CreateField "TblUsers", "USERautoIssueVoucher", adBoolean, adColNullable, , , "", False, True
  
  
  DB_CreateField "TblItemsUnits", "MaxSelingPrice", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblItemsUnits", "SelingPriceDestr", adDouble, adColNullable, , , "    ", False, True




     sql = "DROP FUNCTION GetReyurnedqty"
    Cn.Execute sql
    
  
sql = "  CREATE FUNCTION GetReyurnedqty(@fromdate datetime ,@TOdate datetime  ,@StoreID as integer ,@IDItem as integer )"
sql = sql & "         RETURNS Float            AS Begin" & CHR(13)
sql = sql & "  RETURN ( SELECT        SUM(dbo.Transaction_Details.ShowQty) From dbo.Transaction_Details" & CHR(13)
sql = sql & "            Inner Join               dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
sql = sql & " Where dbo.transactions.Transaction_Type = 9" & CHR(13)
sql = sql & "  AND dbo.Transaction_Details.Item_ID =" & CHR(13)
sql = sql & " ( SELECT        dbo.TblAotherItems.ItemID                From dbo.TblAotherItems" & CHR(13)
sql = sql & "  WHERE        dbo.TblAotherItems.IDItem = @IDItem" & CHR(13)
sql = sql & "AND dbo.Transactions.StoreID = @StoreID" & CHR(13)
sql = sql & " and transactions.Transaction_Date >=@fromdate" & CHR(13)
sql = sql & "  and transactions.Transaction_Date <= @TOdate)" & CHR(13)
sql = sql & " )" & CHR(13)
sql = sql & " End" & CHR(13)
' sql = sql & " End"
 
    db_createOrUpdateFuctionSQL "GetReyurnedqty", sql



DB_CreateField "TblAqrOwin", "ValueAfterDiscount", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAqrOwin", "Discount", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "Transactions", "chkDone", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "Transactions", "DepandToConv", adInteger, adColNullable, , , "    ", False, True

 


    sql = "    DROP FUNCTION FN_MAIN_ACCOUNT_SUB_CODES " & CHR(13)

    Cn.Execute sql
Dim s As String
s = " Create FUNCTION [dbo].[FN_MAIN_ACCOUNT_SUB_CODES]" & CHR(13)
s = s & " ("
s = s & "     @FROMCODE       NVARCHAR(20), " & CHR(13)
s = s & "     @TOCODE         NVARCHAR(20), " & CHR(13)
s = s & "     @WithParent     AS BIT " & CHR(13)
s = s & " )"
s = s & " RETURNS @TABLE TABLE (CODE NVARCHAR(20)) " & CHR(13)
s = s & " AS"


s = s & " Begin " & CHR(13)
s = s & "     IF @WithParent = 0 " & CHR(13)
s = s & "     Begin " & CHR(13)
s = s & "         DECLARE @IsMain BIT " & CHR(13)
s = s & "         DECLARE @Level AS SMALLINT " & CHR(13)
s = s & "         SET @IsMain = @WithParent " & CHR(13)
s = s & "         DECLARE @Child_Code NVARCHAR(20) " & CHR(13)
s = s & "         SET @Child_Code = '' " & CHR(13)
s = s & "         DECLARE @NextLevel SMALLINT " & CHR(13)
s = s & "         DECLARE my_Cursor CURSOR " & CHR(13)
s = s & "         FOR"
s = s & "             SELECT T1.Account_Code, " & CHR(13)
s = s & "                    T1.last_account " & CHR(13)
s = s & "             From Accounts " & CHR(13)
s = s & "                    LEFT OUTER JOIN ( " & CHR(13)
s = s & "                             SELECT Parent_Account_Code, " & CHR(13)
s = s & "                                    Account_Code, " & CHR(13)
s = s & "                                    last_account " & CHR(13)
s = s & "                             From Accounts " & CHR(13)

s = s & "                         )T1 " & CHR(13)
s = s & "                         ON  T1.Parent_Account_Code = ACCOUNTS.Account_Code " & CHR(13)
s = s & "             Where Accounts.last_account = 0 " & CHR(13)

s = s & "                    AND ( " & CHR(13)
s = s & "                            (ACCOUNTS.Account_Code >= @FromCode AND @FromCode <> '') " & CHR(13)
s = s & "                            OR (1 = 1 AND @FromCode = '') " & CHR(13)
s = s & "                        ) " & CHR(13)
s = s & "                    AND ( " & CHR(13)
s = s & "                            (ACCOUNTS.Account_Code <= @ToCode AND @ToCode <> '') " & CHR(13)
s = s & "                            OR (1 = 1 AND @ToCode = '') " & CHR(13)
s = s & "                        ) " & CHR(13)
        
s = s & "         OPEN my_Cursor " & CHR(13)
        
s = s & "         FETCH NEXT FROM my_Cursor INTO @Child_Code,@IsMain " & CHR(13)
        
s = s & "         WHILE @@FETCH_STATUS = 0 " & CHR(13)
         
s = s & "         Begin " & CHR(13)
s = s & "             IF @IsMain = 1 " & CHR(13)
s = s & "             Begin " & CHR(13)
s = s & "                 INSERT INTO @Table " & CHR(13)
s = s & "                 Values " & CHR(13)
s = s & "                   ( " & CHR(13)
s = s & "                     @Child_Code " & CHR(13)
s = s & "                   ) " & CHR(13)
s = s & "             End " & CHR(13)
            
s = s & "             IF @IsMain = 0 " & CHR(13)
s = s & "             Begin " & CHR(13)
s = s & "                 SET @NextLevel = @Level + 1 " & CHR(13)
                
s = s & "                 INSERT @Table " & CHR(13)
s = s & "                 SELECT * " & CHR(13)
s = s & "                 FROM   dbo.fn_Main_Account_Sub_Codes (@Child_Code, @Child_Code, @IsMain) " & CHR(13)
s = s & "             End " & CHR(13)
            
s = s & "             FETCH NEXT FROM my_Cursor INTO @Child_Code,@IsMain " & CHR(13)
s = s & "         End " & CHR(13)
        
s = s & "         Close my_Cursor " & CHR(13)

s = s & "         DEALLOCATE my_Cursor " & CHR(13)

s = s & "         Return " & CHR(13)
s = s & "     End " & CHR(13)
s = s & "     Else " & CHR(13)
s = s & "     Begin " & CHR(13)
s = s & "         DECLARE @TblRet TABLE(Code NVARCHAR(20)) " & CHR(13)
s = s & "         DECLARE @vCurrentNodeCode VARCHAR(50) " & CHR(13)
s = s & "         IF @FROMCODE IS NULL " & CHR(13)
s = s & "            OR @FROMCODE = '' " & CHR(13)
s = s & "             Return " & CHR(13)
        
s = s & "         DECLARE CostCentersCurChildCur CURSOR " & CHR(13)
s = s & "         FOR " & CHR(13)
s = s & "             SELECT Account_Code " & CHR(13)
s = s & "             From [dbo].[Accounts] " & CHR(13)
s = s & "             WHERE  Parent_Account_Code = @FROMCODE " & CHR(13)
        
s = s & "         OPEN CostCentersCurChildCur " & CHR(13)
s = s & "         FETCH NEXT FROM CostCentersCurChildCur " & CHR(13)
s = s & "         INTO @vCurrentNodeCode " & CHR(13)
s = s & "         WHILE @@FETCH_STATUS = 0 " & CHR(13)
s = s & "         Begin " & CHR(13)
s = s & "             INSERT INTO @TABLE " & CHR(13)
s = s & "               ( " & CHR(13)
s = s & "                 code " & CHR(13)
s = s & "               ) " & CHR(13)
s = s & "             Values " & CHR(13)
s = s & "               ( " & CHR(13)
s = s & "                 @vCurrentNodeCode " & CHR(13)
s = s & "               ) " & CHR(13)
s = s & "             INSERT INTO @TABLE " & CHR(13)
s = s & "             SELECT Code " & CHR(13)
s = s & "             FROM   dbo.FN_MAIN_ACCOUNT_SUB_CODES(@vCurrentNodeCode,@vCurrentNodeCode,@WithParent) " & CHR(13)
            
s = s & "             FETCH NEXT FROM CostCentersCurChildCur " & CHR(13)
s = s & "             INTO @vCurrentNodeCode " & CHR(13)
s = s & "         End " & CHR(13)
s = s & "         Close CostCentersCurChildCur " & CHR(13)
s = s & "         DEALLOCATE CostCentersCurChildCur " & CHR(13)
s = s & "     End " & CHR(13)
s = s & "     Return " & CHR(13)
    

s = s & " End " & CHR(13)

 Cn.Execute s
 
 
    New_View = "  SELECT        dbo.Notes.ManualNo, dbo.Notes.OldNoteSerial1, dbo.marakes_taklefa_temp.cost_center_id, dbo.Notes.NoteDateH, dbo.marakes_taklefa_temp.cost_center, dbo.marakes_taklefa_temp.value AS cc_valie, "
  New_View = New_View & "                          dbo.marakes_taklefa_temp.depit_or_credit, dbo.ACCOUNTS.Account_Name, dbo.Notes.NoteSerial, dbo.marakes_taklefa_temp.Project__code, dbo.marakes_taklefa_temp.Project_name, dbo.ACCOUNTS.Account_Serial,"
New_View = New_View & "                            dbo.marakes_taklefa_temp.Description, dbo.Notes.NoteDate, dbo.Notes.Remark, dbo.Notes.RemarkE, dbo.Notes.note_value_by_characters, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.Notes.NoteType, dbo.marakes_taklefa_temp.opr_type, dbo.marakes_taklefa_temp.value AS cc_valie1, dbo.marakes_taklefa_temp.value AS DEV_Value1,"
New_View = New_View & "                            dbo.marakes_taklefa_temp.value AS DEV_Value2, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.Value,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.DOUBLE_ENTREY_VOUCHERS.AdvanceID,"
 New_View = New_View & "                           dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.DOUBLE_ENTREY_VOUCHERS.Posted, dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_Serial, dbo.DOUBLE_ENTREY_VOUCHERS.credit_value, dbo.DOUBLE_ENTREY_VOUCHERS.depet_value,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.des, dbo.DOUBLE_ENTREY_VOUCHERS.currency, dbo.DOUBLE_ENTREY_VOUCHERS.project_bill_no, dbo.DOUBLE_ENTREY_VOUCHERS.valuee, dbo.DOUBLE_ENTREY_VOUCHERS.rate,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.Notes.DueDate, dbo.Notes.NoteHijriDate, dbo.Notes.Transaction_ID AS Expr1, dbo.Notes.MaintananceID,"
 New_View = New_View & "                           dbo.Notes.Member_ID, dbo.Notes.UserID AS Expr2, dbo.Notes.ExpensesID, dbo.Notes.CashingType, dbo.Notes.CusID, dbo.Notes.BoxID, dbo.Notes.RevenuesID, dbo.Notes.RetrunNoteID, dbo.Notes.NoteCashingType,"
 New_View = New_View & "                           dbo.Notes.NotePosted, dbo.Notes.PostedBy, dbo.Notes.PostDate, dbo.Notes.NumOrderInpot, dbo.Notes.ked_type, dbo.Notes.Buy, dbo.Notes.numbering_type, dbo.Notes.sanad_year, dbo.Notes.sanad_month, dbo.Notes.type,"
New_View = New_View & "                           dbo.Notes.branch_no, dbo.Notes.user_name, dbo.Notes.DEPARTEMENT, dbo.Notes.sanad_type, dbo.Notes.sanad_source, dbo.Notes.DAWRY, dbo.Notes.KALEB, dbo.Notes.projectAccountCode, dbo.Notes.foxy_no,"
 New_View = New_View & "                           dbo.Notes.person, dbo.Notes.project_Expensen_account, dbo.Notes.salary, dbo.Notes.displayed, dbo.Notes.Adv_payment_value, dbo.Notes.salary_or_advance, dbo.Notes.EmpAccountCode, dbo.Notes.project_depit_or_credit,"
New_View = New_View & "                            dbo.Notes.Cus_or_sub, dbo.Notes.numbering_type1, dbo.Notes.NoteSerial1, dbo.Notes.general_cost_center, dbo.Notes.too, dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.bill_id, dbo.ACCOUNTS.Account_NameEng,"
New_View = New_View & "                            dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesTypes.NotesTypeNamee,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.Remarks2, dbo.DOUBLE_ENTREY_VOUCHERS.projectid, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.DOUBLE_ENTREY_VOUCHERS.pandid,"
New_View = New_View & "                           dbo.projects.Project_name AS Project_name1, dbo.projects.Fullcode AS ProjectFullcode, dbo.projects.Project_nameE AS Project_nameE1, dbo.Notes.akarid, dbo.Notes.unittype, dbo.Notes.UnitNo, TblAqar_1.aqarname,"
New_View = New_View & "                            TblAqarDetai_1.unitno AS Nameunitno, TblAkarUnit_1.name, TblAkarUnit_1.namee, dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.DOUBLE_ENTREY_VOUCHERS.unittype AS unittype2,"
New_View = New_View & "                            dbo.DOUBLE_ENTREY_VOUCHERS.unitno AS unitno2, TblAqar_1.aqarname AS aqarname2, TblAkarUnit_1.name AS UnitTypename, TblAkarUnit_1.namee AS UnitTypenameE, TblAqarDetai_1.unitno AS Nameunitno2,"
 New_View = New_View & "                           dbo.DOUBLE_ENTREY_VOUCHERS.Departementid , dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee"
New_View = New_View & "   FROM            dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
New_View = New_View & "                            dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code INNER JOIN"
New_View = New_View & "                            dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID INNER JOIN"
New_View = New_View & "                            dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType INNER JOIN"
New_View = New_View & "                            dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
New_View = New_View & "                            dbo.TblEmpDepartments ON dbo.DOUBLE_ENTREY_VOUCHERS.Departementid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
New_View = New_View & "                            dbo.TblAqarDetai AS TblAqarDetai_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unitno = TblAqarDetai_1.Id LEFT OUTER JOIN"
 New_View = New_View & "                           dbo.TblAkarUnit AS TblAkarUnit_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.unittype = TblAkarUnit_1.id LEFT OUTER JOIN"
New_View = New_View & "                            dbo.TblAqar AS TblAqar_1 ON dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid = TblAqar_1.Aqarid LEFT OUTER JOIN"
New_View = New_View & "                            dbo.TblAkarUnit AS TblAkarUnit_2 ON dbo.Notes.unittype = TblAkarUnit_2.id LEFT OUTER JOIN"
 New_View = New_View & "                           dbo.TblAqarDetai AS TblAqarDetai_2 ON dbo.Notes.UnitNo = TblAqarDetai_2.Id LEFT OUTER JOIN"
   New_View = New_View & "                         dbo.TblAqar AS TblAqar_2 ON dbo.Notes.akarid = TblAqar_2.Aqarid LEFT OUTER JOIN"
   New_View = New_View & "                         dbo.projects ON dbo.DOUBLE_ENTREY_VOUCHERS.projectid = dbo.projects.id or dbo.DOUBLE_ENTREY_VOUCHERS.project_id = dbo.projects.id  LEFT OUTER JOIN"
   New_View = New_View & "                         dbo.marakes_taklefa_temp ON dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1 = dbo.marakes_taklefa_temp.line_no AND dbo.marakes_taklefa_temp.line_no <> 0"
                         
    
   ' db_createOrUpdateviewSQL "gl_cc", New_View
    DB_CreateField "project_billl", "CBoBasedON", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "project_billl", "OrDer_no2", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "project_billl", "OrDer_no", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "project_bill_details", "projectName", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_bill_details", "FullCode", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_bill_details", "project_id", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "SubcontractorContract2", "project_id", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "CBoBasedON", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "Notes", "OrDer_no2", adInteger, adColNullable, , , "    ", False, True


     
    'DB_CreateField "Notes", "OrDer_no", adInteger, adColNullable, , , "    ", False, True
'DB_CreateField "Notes", "OrDer_no2", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "tblusers", "CreditLimitSalesMan", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "tblusers", "CreditLimitSalesMan", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "tblusers", "NotEditInternalPrice", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "tblusers", "NotEditSalesRetPrice", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "SubcontractorContract2", "Pre_Quantity_Contr", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "project_bill_details", "Pre_Quantity_Contr", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "SubcontractorContract2", "qtySubContractor", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "project_bill_details", "qtySubContractor", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblAging", "NoteType", adInteger, adColNullable, , , "                ", False, True


DB_CreateField "TblAging", "NoteType", adInteger, adColNullable, , , "                ", False, True



DB_CreateField "TblOptions", "IsInternalMultiOrder", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "TblVATAvowal", "tztAdvBill", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Groups", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "tblUsers", "IsDiscountPerLine", adBoolean, adColNullable, , , "                ", False, True

  DB_CreateField "Transactions", "TotalTaxExempt", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblusers", "NotEditDiscountLine", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "tblusers", "CanEditMinRentValue", adBoolean, adColNullable, , , "                ", False, True



DB_CreateField "Groups", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "Transactions", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "tblUsers", "IsDiscountPerLine", adBoolean, adColNullable, , , "                ", False, True

  DB_CreateField "Transactions", "TotalTaxExempt", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblusers", "NotEditDiscountLine", adBoolean, adColNullable, , , "                ", False, True

 

DB_CreateField "Groups", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "Transactions", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "tblUsers", "IsDiscountPerLine", adBoolean, adColNullable, , , "                ", False, True

  DB_CreateField "Transactions", "TotalTaxExempt", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "tblusers", "NotEditDiscountLine", adBoolean, adColNullable, , , "                ", False, True


DB_CreateField "TblVATAvowal", "SaleWithOutVat", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "RetSaleWithOutVat", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "TblVATAvowal", "tztAdvBill", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblVATAvowal", "ChkIsFree", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "TblAqrOwin", "Select", adBoolean, adColNullable, , , "                ", False, True

DB_CreateField "TblAqrOwin", "valueBeforDiscount", adDouble, adColNullable, , , "    ", False, True



If DB_CreateTable("TblContractInstallDisco2", True, "id ", True) = True Then
    DB_CreateField "TblContractInstallDisco2", "valuewithout", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "VatPerc", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "VatValue", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblContractInstallDisco2", "ValueAfterDiscount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "Discount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "DiscountValue", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "Value", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco2", "Select", adBoolean, adColNullable, , , "                ", False, True
    DB_CreateField "TblContractInstallDisco2", "PaymentNo", adInteger, adColNullable, , , "                ", False, True
    DB_CreateField "TblContractInstallDisco2", "MasterNo", adInteger, adColNullable, , , "                ", False, True
    DB_CreateField "TblContractInstallDisco2", "MasterID", adInteger, adColNullable, , , "                ", False, True
    DB_CreateField "TblContractInstallDisco2", "Cont", adInteger, adColNullable, , , "                ", False, True
    DB_CreateField "TblContractInstallDisco2", "Ser", adInteger, adColNullable, , , "                ", False, True
    
    DB_CreateField "TblContractInstallDisco2", "RecDateH", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblContractInstallDisco2", "AllowDateH", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    DB_CreateField "TblContractInstallDisco2", "DMY", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    
    
    
    
    DB_CreateField "TblContractInstallDisco2", "AllowDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblContractInstallDisco2", "RecDate", adDBTimeStamp, adColNullable, , , "", False, True

End If


    
        
If DB_CreateTable("TblContractInstallDisco", True, "id ", False) = True Then
    DB_CreateField "TblContractInstallDisco", "RecordDate", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblContractInstallDisco", "DiscountVal", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco", "InstallNoStart", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco", "DiscountType", adInteger, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblContractInstallDisco", "UserID", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco", "CusId", adInteger, adColNullable, , , "    ", False, True
    DB_CreateField "TblContractInstallDisco", "Iqar", adInteger, adColNullable, , , "    ", False, True

End If


End Function



Public Function projectincludevchr()
Dim str As String
Dim s As String
s = "  update Notes   Set OrderIDD = Noteseril2"
s = s & "  FROM         dbo.Notes INNER JOIN"
s = s & "                        dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
s = s & "  Where (dbo.notes_all.NoteType = 3)"
'Cn.Execute s


s = " update TransactionTypes set projectInclude=1  where  (Transaction_Type=66 OR Transaction_Type=18)"
Cn.Execute s

End Function


Public Function ApprovalScreen()
Dim str As String

str = "update Screens Set FlgShow = null"
Cn.Execute str

str = "update Screens "
str = str & " Set FlgShow = 1"
str = str & " where ScreenName='FrmAccEditJournal'"
str = str & " or ScreenName='FrmAccEditJournal1'"
str = str & " or ScreenName='FrmAccEditJournal3'"
str = str & " or ScreenName='FrmAccEditJournal4'"
str = str & " or ScreenName='FrmExpenses3'"
str = str & " or ScreenName='FrmExpenses5'"
str = str & " or ScreenName='FrmTypeExchange'"
str = str & " or ScreenName='FrmCashing'"
str = str & " or ScreenName='FrmPayments1'"
str = str & " or ScreenName='FrmExpenses30'"
str = str & " or ScreenName='FrmBoxDrawing'"
str = str & " or ScreenName='FrmExpenses4'"
str = str & " or ScreenName='formvocatinl'"
str = str & " or ScreenName='FrmEmpsAdvanceRequest'"
str = str & " or ScreenName='FrmPO1'"
str = str & " or ScreenName='FrmPO2'"
str = str & " or ScreenName='FrmPO3'"
str = str & " or ScreenName='FrmPO4'"
str = str & " or ScreenName='FrmPO5'"
str = str & " or ScreenName='FrmPO6'"
str = str & " or ScreenName='FrmPO7'"
str = str & " or ScreenName='FrmPO8'"
str = str & " or ScreenName='FrmPO11'"
str = str & " or ScreenName='FrmPO10'"
str = str & " or ScreenName='FrmDestruction'"
str = str & " or ScreenName='FrmDestructionRet'"
str = str & " or ScreenName='FrmMovingEmp'"
str = str & " or ScreenName='formempmovedepartment'"
str = str & " or ScreenName='FrmAdvancedHousingpayments'"
str = str & " or ScreenName='FrmQuesEmp'"
str = str & " or ScreenName='FrmBankPledge1' "
str = str & " or ScreenName='FrmBankPledge2' "
str = str & " or ScreenName='FrmBankPledge3' "
str = str & " or ScreenName='FrmBankPledge4' "
str = str & " or ScreenName='FrmExpenses301' "
str = str & " or ScreenName='End_oF_service' "
str = str & " or ScreenName='FrmVocationEntitlements' "
 str = str & " or ScreenName='FrmMoving' "
 str = str & " or ScreenName='projectsbill' "
 

 
Cn.Execute str

str = " delete Screens WHERE        (ScreenType = 20) AND (ScreenName = N'FrmPO10')"

Cn.Execute str

End Function
 
 
Public Function chek()
 
        



End Function
 

Public Function updateExpenses_Oreder_qry()
 

End Function

Public Function GardByUnits()
 

End Function

Public Function AgingReports()
 
End Function

Public Function Create_items_status_report(Optional runreport As Integer = 0, _
                                           Optional StrSQLDate As String)
    Dim New_View  As String
    '20 12 2011

    New_View = " SELECT DISTINCT order_no, MAX(OrderArrivalDate) AS OrderArrivalDateMax"
    New_View = New_View & "  From dbo.Transaction_Details "
 
    New_View = New_View & "   GROUP BY order_no"
    New_View = New_View & "   Having (Not (order_no Is Null)) And (Not (Max(OrderArrivalDate) Is Null))"
    db_createOrUpdateviewSQL "QRyOrdersData", New_View

    New_View = " SELECT     TOP 100 PERCENT SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity * dbo.TblItems.SallingPrice) AS qty, "
    New_View = New_View & "   SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.Quantity) AS Realqty, dbo.Transaction_Details.order_no"
    New_View = New_View & "   FROM         dbo.Transaction_Details INNER JOIN"

    New_View = New_View & "    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    New_View = New_View & "    dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    New_View = New_View & "    dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "

    If runreport = 1 Then
        New_View = New_View & StrSQLDate
    End If

    New_View = New_View & "   GROUP BY dbo.Transaction_Details.order_no"

    db_createOrUpdateviewSQL "Items_Status1", New_View
 
    New_View = "SELECT     dbo.Transaction_Details.order_no, dbo.Transactions.StoreID, dbo.TblStore.StoreName, SUM(dbo.Transaction_Details.OpeningBurcahseQty) AS totalPurchaseQty, "
    New_View = New_View & "  SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS actualQty, SUM(dbo.Transaction_Details.OpeningBurcahseValue)"
    New_View = New_View & "   AS totalPurchaseValue, SUM(dbo.Transaction_Details.OpeningSalesQty) AS totalSalesQty, SUM(dbo.Transaction_Details.OpeningSalesValue) AS SaleValueValue,"
    New_View = New_View & "  dbo.Items_Status1.qty AS SaleValue1"
    New_View = New_View & "  FROM         dbo.Transactions INNER JOIN"
    New_View = New_View & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    New_View = New_View & "   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    New_View = New_View & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
    New_View = New_View & "   dbo.Items_Status1 ON dbo.Transaction_Details.order_no = dbo.Items_Status1.order_no"

    If runreport = 1 Then
        New_View = New_View & StrSQLDate
    End If

    New_View = New_View & "  GROUP BY dbo.Transaction_Details.order_no, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Items_Status1.qty"
    db_createOrUpdateviewSQL "Items_Status2", New_View

    If runreport = 1 Then
        Exit Function
    End If

    New_View = " SELECT     order_no AS order_nox1, StoreID AS StoreIDx1, StoreName AS StoreNamex1, totalPurchaseQty AS totalPurchaseQtyx1, totalPurchaseValue AS totalPurchaseValuex1, "
    New_View = New_View & "   totalSalesQty As totalSalesQtyx1, SaleValueValue As SaleValueValuex1, SaleValue1 As SaleValuex1, actualQty As actualQtyx1"
    New_View = New_View & "   From dbo.Items_Status2 Where (StoreID = 1)"
    db_createOrUpdateviewSQL "Items_Status_totals1", New_View

    New_View = "SELECT     order_no AS order_nox2, StoreID AS StoreIDx2, StoreName AS StoreNamex2, totalPurchaseQty AS totalPurchaseQtyx2, totalPurchaseValue AS totalPurchaseValuex2, "
    New_View = New_View & "  totalSalesQty As totalSalesQtyx2, SaleValueValue As SaleValueValuex2, SaleValue1 As SaleValuex2, actualQty As actualQtyx2"
    New_View = New_View & "   From dbo.Items_Status2"
    New_View = New_View & "   WHERE     (StoreID = 2)"
    db_createOrUpdateviewSQL "Items_Status_totals2", New_View

    New_View = "SELECT     order_no AS order_nox3, StoreID AS StoreIDx3, StoreName AS StoreNamex3, totalPurchaseQty AS totalPurchaseQtyx3, totalPurchaseValue AS totalPurchaseValuex3, "
    New_View = New_View & "   totalSalesQty As totalSalesQtyx3, SaleValueValue As SaleValueValuex3, SaleValue1 As SaleValuex3, actualQty As actualQtyx3"
    New_View = New_View & "   From dbo.Items_Status2"
    New_View = New_View & "   WHERE     (StoreID = 3)"
    db_createOrUpdateviewSQL "Items_Status_totals3", New_View

    New_View = "SELECT     order_no AS order_nox4, StoreID AS StoreIDx4, StoreName AS StoreNamex4, totalPurchaseQty AS totalPurchaseQtyx4, totalPurchaseValue AS totalPurchaseValuex4, "
    New_View = New_View & "  totalSalesQty As totalSalesQtyx4, SaleValueValue As SaleValueValuex4, SaleValue1 As SaleValuex4, actualQty As actualQtyx4"
    New_View = New_View & " From dbo.Items_Status2"
    New_View = New_View & "  WHERE     (StoreID = 4)"
    db_createOrUpdateviewSQL "Items_Status_totals4", New_View

    New_View = "SELECT     order_no AS order_nox5, StoreID AS StoreIDx5, StoreName AS StoreNamex5, totalPurchaseQty AS totalPurchaseQtyx5, totalPurchaseValue AS totalPurchaseValuex5, "
    New_View = New_View & " totalSalesQty As totalSalesQtyx5, SaleValueValue As SaleValueValuex5, SaleValue1 As SaleValuex5, actualQty As actualQtyx5"
    New_View = New_View & " From dbo.Items_Status2"
    New_View = New_View & "  WHERE     (StoreID = 5)"
    db_createOrUpdateviewSQL "Items_Status_totals5", New_View

    New_View = "SELECT     order_no AS order_nox6, StoreID AS StoreIDx6, StoreName AS StoreNamex6, totalPurchaseQty AS totalPurchaseQtyx6, totalPurchaseValue AS totalPurchaseValuex6, "
    New_View = New_View & " totalSalesQty As totalSalesQtyx6, SaleValueValue As SaleValueValuex6, SaleValue1 As SaleValuex6, actualQty As actualQtyx6"
    New_View = New_View & " From dbo.Items_Status2"
    New_View = New_View & "  WHERE     (StoreID = 6)"
    db_createOrUpdateviewSQL "Items_Status_totals6", New_View

    New_View = "SELECT     order_no AS order_nox7, StoreID AS StoreIDx7, StoreName AS StoreNamex7, totalPurchaseQty AS totalPurchaseQtyx7, totalPurchaseValue AS totalPurchaseValuex7, "
    New_View = New_View & "   totalSalesQty As totalSalesQtyx7, SaleValueValue As SaleValueValuex7, SaleValue1 As SaleValuex7, actualQty As actualQtyx7"
    New_View = New_View & " From dbo.Items_Status2"
    New_View = New_View & "  WHERE     (StoreID = 7)"

    db_createOrUpdateviewSQL "Items_Status_totals7", New_View

End Function
Public Function UpdateCostriceProcedureByStores()
On Error Resume Next
    sql = "    DROP FUNCTION QryItemsTransactionsTotalsByStores" & CHR(13)
    Cn.Execute sql

    sql = " CREATE FUNCTION QryItemsTransactionsTotalsByStores(@TransType int =0,@TransType2 int=0,@TransType3 int=0,@FromDate datetime ,@ToDate datetime ,@storeid as integer,@ItemID  as integer,@Transaction_ID as float=null )" & CHR(13)
   sql = sql & "  RETURNS @xTable TABLE" & CHR(13)
   sql = sql & " (" & CHR(13)
   sql = sql & " ItemID int," & CHR(13)
   sql = sql & " ItemCode nvarchar(50)," & CHR(13)
   sql = sql & "  ItemName nvarchar(4000)," & CHR(13)
sql = sql & "  GroupID  int," & CHR(13)
sql = sql & "  Total   money," & CHR(13)
sql = sql & "  totalqty Float" & CHR(13)
sql = sql & "  )" & CHR(13)
sql = sql & "  AS" & CHR(13)
sql = sql & "  Begin" & CHR(13)
sql = sql & "  INSERT @xTable" & CHR(13)
   sql = sql & "  Select ItemID,ItemCode,ItemName,GroupID,Sum(Total) as Totals,Sum(Quantity) as TotalQty" & CHR(13)
sql = sql & "  From" & CHR(13)
sql = sql & "  (" & CHR(13)
sql = sql & "  SELECT TblItems.ItemID,TblItems.ItemCode, TblItems.ItemName,TblItems.GroupID," & CHR(13)
sql = sql & "   'Total'=Case" & CHR(13)
sql = sql & "  When ItemDiscountType=1 Or ItemDiscountType=0 Then Transaction_Details.Quantity*Transaction_Details.Price" & CHR(13)
sql = sql & "  When ItemDiscountType=2 Then ((Transaction_Details.Quantity*Transaction_Details.Price)-ItemDiscount)" & CHR(13)
sql = sql & "  When ItemDiscountType=3 Then (Transaction_Details.Quantity*Transaction_Details.Price) *( 1- (ItemDiscount/100))" & CHR(13)
sql = sql & "  Else  0" & CHR(13)
sql = sql & "  End" & CHR(13)
sql = sql & "  ,Transaction_Details.Quantity " & CHR(13)
sql = sql & "  FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN" & CHR(13)
sql = sql & "  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
sql = sql & "  WHERE (Transactions.Transaction_Type=@TransType  OR Transactions.Transaction_Type=@TransType2 OR Transactions.Transaction_Type=@TransType3" & CHR(13)
sql = sql & "  or   Transactions.Transaction_Type=34 or   Transactions.Transaction_Type=15   )" & CHR(13)
sql = sql & "  AND" & CHR(13)
sql = sql & "  Transactions.Transaction_Date >=@FromDate" & CHR(13)
sql = sql & "  AND" & CHR(13)
sql = sql & "  Transactions.Transaction_Date <=@TODate" & CHR(13)
sql = sql & "  AND" & CHR(13)
sql = sql & "  Transactions.storeid =@storeid" & CHR(13)
sql = sql & " and  Transactions.Transaction_ID<>isnull(@Transaction_ID,Transactions.Transaction_ID)" & CHR(13)
 
' sql = sql & " and  Transaction_Details.UnitId = ISNULL(@UnitID,Transaction_Details.UnitId)" & CHR(13)
sql = sql & "  AND" & CHR(13)
sql = sql & "  TblItems.ItemID =@ItemID" & CHR(13)


sql = sql & "  )DrivTable" & CHR(13)
sql = sql & "  Group By ItemID,ItemCode,ItemName,GroupID" & CHR(13)
sql = sql & "  Return" & CHR(13)
 sql = sql & "  End" & CHR(13)



 
    db_createOrUpdateFuctionSQL "QryItemsTransactionsTotalsByStores", sql





 s = " create  FUNCTION [dbo].[GetBalanceQtyPO5] (@ItemID integer ,@order_no  nvarchar(255) ,@PurchaseNo  integer,@TransType  integer,@CBoBasedON  integer )  RETURNS Float AS Begin"
 s = s & " Return (SELECT"
s = s & "         SUM(dbo.Transaction_Details.ShowQty) As ShowQty"
s = s & "     From dbo.Transaction_Details"
s = s & "     RIGHT OUTER JOIN dbo.Transactions"
s = s & "         ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
s = s & "     Where dbo.Transactions.Transaction_Type = 29"
s = s & "     AND (dbo.Transactions.NoteSerial1 = @order_no)"
s = s & "     AND (dbo.Transaction_Details.Item_ID = @ItemID))"
s = s & " - ISNULL((SELECT"
s = s & "         SUM(dbo.Transaction_Details.ShowQty) As ShowQty"
s = s & "     From dbo.Transaction_Details"
s = s & "     RIGHT OUTER JOIN dbo.Transactions"
s = s & "         ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
s = s & "     WHERE dbo.transactions.Transaction_Type = @TransType"
s = s & "     AND (dbo.Transactions.order_no = @order_no)"
s = s & "     AND (ISNULL(dbo.Transactions.CBoBasedON, 0) = @CBoBasedON)"
s = s & "     AND (dbo.Transaction_Details.Item_ID = @ItemID"
s = s & "     AND Transactions.Transaction_ID <> @PurchaseNo))"
s = s & " , 0)"
s = s & " End"
db_createOrUpdateFuctionSQL "GetBalanceQtyPO5", s


 s = " create  FUNCTION [dbo].[GetBalanceQtyPO4] (@ItemID integer ,@order_no  nvarchar(255) ,@PurchaseNo  integer,@TransType  integer,@CBoBasedON  integer )  RETURNS Float AS Begin"
 s = s & " Return (SELECT"
s = s & "         SUM(dbo.Transaction_Details.ShowQty) As ShowQty"
s = s & "     From dbo.Transaction_Details"
s = s & "     RIGHT OUTER JOIN dbo.Transactions"
s = s & "         ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
s = s & "     Where dbo.Transactions.Transaction_Type = 29"
s = s & "     AND (dbo.Transactions.NoteSerial1 = @order_no)"
s = s & "     AND (dbo.Transaction_Details.Item_ID = @ItemID))"
s = s & " - ISNULL((SELECT"
s = s & "         SUM(dbo.Transaction_Details.ShowQty) As ShowQty"
s = s & "     From dbo.Transaction_Details"
s = s & "     RIGHT OUTER JOIN dbo.Transactions"
s = s & "         ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
s = s & "     WHERE dbo.transactions.Transaction_Type = @TransType"
s = s & "     AND (dbo.Transactions.PONo = @order_no)"
s = s & "     AND (ISNULL(dbo.Transactions.CBoBasedON, 0) = @CBoBasedON)"
s = s & "     AND (dbo.Transaction_Details.Item_ID = @ItemID"
s = s & "     AND Transactions.Transaction_ID <> @PurchaseNo))"
s = s & " , 0)"
s = s & " End"

db_createOrUpdateFuctionSQL "GetBalanceQtyPO4", s

s = "CREATE FUNCTION dbo.GetDaysInMonth2 (@Year INT, @Month INT) RETURNS INT "
s = s & " AS"
s = s & " Begin"
s = s & "     DECLARE @FirstDayOfMonth DATE = DATEFROMPARTS(@Year, @Month, 1)"
s = s & "     DECLARE @LastDayOfMonth DATE = EOMONTH(@FirstDayOfMonth)"
s = s & "     RETURN DATEDIFF(DAY, @FirstDayOfMonth, @LastDayOfMonth) + 1"
s = s & " End"


db_createOrUpdateFuctionSQL "GetDaysInMonth2", s

    s = " Create  FUNCTION [dbo].[GetBalanceQtyPO3] (@ItemID integer ,@order_no  nvarchar(255) ,@PurchaseNo  integer )"
    s = s & "  RETURNS Float"
    s = s & " AS"
    s = s & " Begin"
    s = s & " Return"
    
    s = s & " (SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ShowQty"
    s = s & "    FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
    s = s & "                  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "     Where dbo.transactions.Transaction_Type = 29"
    s = s & "                    AND (dbo.Transactions.NoteSerial1 = @order_no)  AND"
    s = s & "                   (dbo.Transaction_Details.Item_ID = @ItemID)  ) -"
    
    s = s & "  IsNull((SELECT     SUM(dbo.Transaction_Details.ShowQty) AS ShowQty"
    s = s & "    FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
    s = s & "                   dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    s = s & "     Where dbo.transactions.Transaction_Type = 22"
    s = s & "                    AND (dbo.Transactions.order_no = @order_no) AND (ISNULL(dbo.Transactions.CBoBasedON, 0) = 1) AND"
    s = s & "                   (dbo.Transaction_Details.Item_ID = @ItemID and Transactions.Transaction_ID <> @PurchaseNo)"
    s = s & "    )  ,0)"
    
    
    s = s & " End"

    db_createOrUpdateFuctionSQL "GetBalanceQtyPO3", s
    
    
    s = "CREATE FUNCTION dbo.GetDaysInMonth (@Year INT, @Month INT) RETURNS INT "
s = s & " AS"
s = s & " Begin"
s = s & "     DECLARE @FirstDayOfMonth DATE = DATEFROMPARTS(@Year, @Month, 1)"
s = s & "     DECLARE @LastDayOfMonth DATE = EOMONTH(@FirstDayOfMonth)"
s = s & "     RETURN DATEDIFF(DAY, @FirstDayOfMonth, @LastDayOfMonth) + 1"
s = s & " End"


db_createOrUpdateFuctionSQL "GetDaysInMonth", s


   s = ""
   s = s & " Create FUNCTION [dbo].[GetAbcentDay2](@EmpID  integer,@YearID  integer,@MonthID  integer )"
   s = s & "    RETURNS Float"
    s = s & "      AS"
    s = s & "     Begin"
    s = s & "     RETURN (    SELECT     SUM(dbo.TblChangedComponentRegisterDetails.NoofDays) AS SumNoofDays"
    s = s & " FROM         dbo.TblChangedComponentRegister LEFT OUTER JOIN"
    s = s & "                    dbo.TblChangedComponentRegisterDetails ON"
    s = s & "                    dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid"
    s = s & "   WHERE      (dbo.TblChangedComponentRegister.Actualmonth = @MonthID) AND"
    s = s & "                     (dbo.TblChangedComponentRegister.Actualyear = @YearID)"
    s = s & "                     AND IsNull(value,0) = 0"
    s = s & "    GROUP BY dbo.TblChangedComponentRegisterDetails.Emp_id"
    s = s & "   Having (dbo.TblChangedComponentRegisterDetails.Emp_id = @EmpID)"
    s = s & " )"
    s = s & "     End"
  db_createOrUpdateFuctionSQL "GetAbcentDay2", s

End Function

Public Function UpdateEmpVoCation2()


Dim StrSQL As String
On Error Resume Next
    StrSQL = "    DROP FUNCTION EmpVoCation2" & CHR(13)
    Cn.Execute StrSQL
    
StrSQL = "CREATE FUNCTION [dbo].[EmpVoCation2] " & _
         "(@Monthh AS INTEGER, @Yar AS INTEGER, @EmpID AS INTEGER) " & _
         "RETURNS FLOAT " & _
         "AS " & _
         "BEGIN " & _
         "    RETURN ( " & _
         "        SELECT SUM(SickDays) AS SickBalance " & _
         "        FROM TblRegsterSickleave2 " & _
         "        WHERE (EmpID = @EmpID) AND (MonthnO = @Monthh) AND (SickYear = @Yar) " & _
         "    ) " & _
         "END"

Cn.Execute StrSQL
End Function
Public Function UpdateAccountProc()
Dim sql As String
sql = sql & "ALTER FUNCTION [dbo].[GetBalance](" & vbCrLf
sql = sql & "    @fromdate     DATETIME," & vbCrLf
sql = sql & "    @Todate       DATETIME," & vbCrLf
sql = sql & "    @accountcode  VARCHAR(255)," & vbCrLf
sql = sql & "    @LastAccount  INT," & vbCrLf
sql = sql & "    @IsHiddenInv  BIT = 0" & vbCrLf
sql = sql & ")" & vbCrLf
sql = sql & "RETURNS FLOAT" & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "    RETURN (" & vbCrLf
sql = sql & "        SELECT SUM(DEV_Value1) - SUM(DEV_Value2) AS result" & vbCrLf
sql = sql & "        FROM (" & vbCrLf
sql = sql & "            SELECT" & vbCrLf
sql = sql & "                DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END," & vbCrLf
sql = sql & "                DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * 1 ELSE 0 END" & vbCrLf
sql = sql & "            FROM dbo.DOUBLE_ENTREY_VOUCHERS" & vbCrLf
sql = sql & "            WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code LIKE" & vbCrLf
sql = sql & "                  CASE WHEN @LastAccount = 1 THEN @accountcode ELSE @accountcode + 'a%' END" & vbCrLf
sql = sql & "              AND (dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >= @fromdate" & vbCrLf
sql = sql & "                   AND dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <= @Todate)" & vbCrLf
sql = sql & "              AND dbo.DOUBLE_ENTREY_VOUCHERS.Posted IS NULL" & vbCrLf
sql = sql & "              AND (@IsHiddenInv = 1 OR IsNull(dbo.DOUBLE_ENTREY_VOUCHERS.IsHiddenInv,0) = 0)" & vbCrLf
sql = sql & "        ) XTABLE" & vbCrLf
sql = sql & "    )" & vbCrLf
sql = sql & "END" & vbCrLf

Cn.Execute sql
Dim MySQL As String
MySQL = MySQL & "SELECT dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Value AS DEV_Value, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, "
MySQL = MySQL & "dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, "
MySQL = MySQL & "dbo.ACCOUNTS.Branch, dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center, dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters, dbo.Notes.foxy_no, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.TblNotesTypes.NotesTypeNamee, dbo.TransactionTypes.TransactionEnglishName, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, "
MySQL = MySQL & "dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.DOUBLE_ENTREY_VOUCHERS.Posted, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.valuee AS DEV_ValueE, dbo.DOUBLE_ENTREY_VOUCHERS.currency, dbo.DOUBLE_ENTREY_VOUCHERS.rate, dbo.TblBranchesData.RegionID, dbo.TblSection.name, dbo.TblSection.namee, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.DescAccount, dbo.DOUBLE_ENTREY_VOUCHERS.NextAccount_Code, dbo.DOUBLE_ENTREY_VOUCHERS.project_id, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS.projectid,DOUBLE_ENTREY_VOUCHERS.IsHiddenInv, dbo.DOUBLE_ENTREY_VOUCHERS.operid, dbo.DOUBLE_ENTREY_VOUCHERS.pandid, dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid, dbo.TblAqar.aqarname, "
MySQL = MySQL & "dbo.TblAqar.aqarNo "
MySQL = MySQL & "FROM dbo.TblAqar RIGHT OUTER JOIN "
MySQL = MySQL & "dbo.TblBranchesData INNER JOIN "
MySQL = MySQL & "dbo.TblUsers INNER JOIN "
MySQL = MySQL & "dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id ON "
MySQL = MySQL & "dbo.TblAqar.Aqarid = dbo.DOUBLE_ENTREY_VOUCHERS.Aqarid LEFT OUTER JOIN "
MySQL = MySQL & "dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN "
MySQL = MySQL & "dbo.TblSection ON dbo.TblBranchesData.RegionID = dbo.TblSection.Id LEFT OUTER JOIN "
MySQL = MySQL & "dbo.Notes LEFT OUTER JOIN "
MySQL = MySQL & "dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN "
MySQL = MySQL & "dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN "
MySQL = MySQL & "dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
MySQL = MySQL & "WHERE (dbo.DOUBLE_ENTREY_VOUCHERS.Posted IS NULL)"


 db_createOrUpdateviewSQL "RptLedger_Sub", MySQL

End Function
Public Function UpdateCostriceProcedure()
On Error Resume Next
   DB_CreateField "TblUsers", "ReportName", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TblUsers", "ReportName1", adVarWChar, adColNullable, 255, , "", False, True, , True
        
    DB_CreateField "TblContractReVouch", "DateK2", adDBTimeStamp, adColNullable, , , "", False, True
    DB_CreateField "TblContractReVouch", "DateK", adDBTimeStamp, adColNullable, , , "", False, True



Dim sql As String
sql = ""


sql = ""
'sql = sql & "IF OBJECT_ID('dbo.QryItemsInventry4', 'IF') IS NOT NULL DROP FUNCTION dbo.QryItemsInventry4" & vbCrLf

sql = sql & "CREATE FUNCTION [dbo].QryItemsInventry4 (" & vbCrLf
sql = sql & "    @fromdate datetime," & vbCrLf
sql = sql & "    @todate datetime," & vbCrLf
sql = sql & "    @StoreId AS INT = null," & vbCrLf
sql = sql & "    @ColorID AS INT = null," & vbCrLf
sql = sql & "    @ItemSize AS NVARCHAR(255) = null," & vbCrLf
sql = sql & "    @ClassId AS INT = null," & vbCrLf
sql = sql & "    @LotNO AS NVARCHAR(255) = null," & vbCrLf
sql = sql & "    @CusID as float = null" & vbCrLf
sql = sql & ") " & vbCrLf
sql = sql & "RETURNS @XTable TABLE (" & vbCrLf
sql = sql & "    Item_ID DECIMAL(18,2)," & vbCrLf
sql = sql & "    ColorID INT," & vbCrLf
sql = sql & "    ItemSize NVARCHAR(255)," & vbCrLf
sql = sql & "    ClassId INT," & vbCrLf
sql = sql & "    LotNO NVARCHAR(255)," & vbCrLf
sql = sql & "    UnitId INT," & vbCrLf
sql = sql & "    ItemCode NVARCHAR(255)," & vbCrLf
sql = sql & "    ItemName NVARCHAR(4000)," & vbCrLf
'sql = sql & "    Length DECIMAL(18,2) NULL," & vbCrLf
'sql = sql & "    Height DECIMAL(18,2) NULL," & vbCrLf
'sql = sql & "    Width DECIMAL(18,2) NULL," & vbCrLf
sql = sql & "    openingValue DECIMAL(18,2)," & vbCrLf
sql = sql & "    inputvalue DECIMAL(18,2)," & vbCrLf
sql = sql & "    outputValue DECIMAL(18,2)," & vbCrLf
sql = sql & "    openingBalance DECIMAL(18,2)" & vbCrLf
sql = sql & ") " & vbCrLf
sql = sql & "AS" & vbCrLf
sql = sql & "BEGIN" & vbCrLf
sql = sql & "    ;WITH AllComb AS (" & vbCrLf
sql = sql & "        SELECT DISTINCT" & vbCrLf
sql = sql & "            TD.Item_ID," & vbCrLf
sql = sql & "            ISNULL(TD.ColorID, 0) AS ColorID," & vbCrLf
sql = sql & "            ISNULL(TD.ItemSize, '') AS ItemSize," & vbCrLf
sql = sql & "            ISNULL(TD.ClassId, 0) AS ClassId," & vbCrLf
sql = sql & "            ISNULL(TD.LotNO, '') AS LotNO," & vbCrLf
sql = sql & "            ISNULL(TD.UnitId, 0) AS UnitId" & vbCrLf
'sql = sql & "            ISNULL(TD.Length, 0) AS Length," & vbCrLf
'sql = sql & "            ISNULL(TD.Height, 0) AS Height," & vbCrLf
'sql = sql & "            ISNULL(TD.Width, 0) AS Width" & vbCrLf
sql = sql & "        FROM dbo.Transaction_Details TD" & vbCrLf
sql = sql & "        INNER JOIN dbo.Transactions T ON TD.Transaction_ID = T.Transaction_ID" & vbCrLf
sql = sql & "        INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type" & vbCrLf
sql = sql & "        WHERE" & vbCrLf
sql = sql & "            T.Transaction_Date <= @todate" & vbCrLf
sql = sql & "            AND (T.Storeid = @StoreId OR @StoreId IS NULL)" & vbCrLf
sql = sql & "            AND TT.StockEffect <> 0" & vbCrLf
sql = sql & "    )," & vbCrLf
sql = sql & "    Opening AS (" & vbCrLf
sql = sql & "        SELECT" & vbCrLf
sql = sql & "            TD.Item_ID," & vbCrLf
sql = sql & "            ISNULL(TD.ColorID, 0) AS ColorID," & vbCrLf
sql = sql & "            ISNULL(TD.ItemSize, '') AS ItemSize," & vbCrLf
sql = sql & "            ISNULL(TD.ClassId, 0) AS ClassId," & vbCrLf
sql = sql & "            ISNULL(TD.LotNO, '') AS LotNO," & vbCrLf
sql = sql & "            ISNULL(TD.UnitId, 0) AS UnitId," & vbCrLf
'sql = sql & "            ISNULL(TD.Length, 0) AS Length," & vbCrLf
'sql = sql & "            ISNULL(TD.Height, 0) AS Height," & vbCrLf
'sql = sql & "            ISNULL(TD.Width, 0) AS Width," & vbCrLf
sql = sql & "            SUM(TD.Quantity * TT.StockEffect) AS openingBalance" & vbCrLf
sql = sql & "        FROM dbo.Transaction_Details TD" & vbCrLf
sql = sql & "        INNER JOIN dbo.Transactions T ON TD.Transaction_ID = T.Transaction_ID" & vbCrLf
sql = sql & "        INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type" & vbCrLf
sql = sql & "        WHERE T.Transaction_Date < @fromdate" & vbCrLf
sql = sql & "        AND (T.Storeid = @StoreId OR @StoreId IS NULL)" & vbCrLf
sql = sql & "        AND TT.StockEffect <> 0" & vbCrLf
'sql = sql & "        GROUP BY TD.Item_ID, ISNULL(TD.ColorID, 0), ISNULL(TD.ItemSize, ''), ISNULL(TD.ClassId, 0), ISNULL(TD.LotNO, ''), ISNULL(TD.UnitId, 0), ISNULL(TD.Length, 0), ISNULL(TD.Height, 0), ISNULL(TD.Width, 0)" & vbCrLf
sql = sql & "        GROUP BY TD.Item_ID, ISNULL(TD.ColorID, 0), ISNULL(TD.ItemSize, ''), ISNULL(TD.ClassId, 0), ISNULL(TD.LotNO, ''), ISNULL(TD.UnitId, 0)" & vbCrLf
sql = sql & "    )," & vbCrLf
sql = sql & "    TransInPeriod AS (" & vbCrLf
sql = sql & "        SELECT" & vbCrLf
sql = sql & "            TD.Item_ID," & vbCrLf
sql = sql & "            ISNULL(TD.ColorID, 0) AS ColorID," & vbCrLf
sql = sql & "            ISNULL(TD.ItemSize, '') AS ItemSize," & vbCrLf
sql = sql & "            ISNULL(TD.ClassId, 0) AS ClassId," & vbCrLf
sql = sql & "            ISNULL(TD.LotNO, '') AS LotNO," & vbCrLf
sql = sql & "            ISNULL(TD.UnitId, 0) AS UnitId," & vbCrLf
'sql = sql & "            ISNULL(TD.Length, 0) AS Length," & vbCrLf
'sql = sql & "            ISNULL(TD.Height, 0) AS Height," & vbCrLf
'sql = sql & "            ISNULL(TD.Width, 0) AS Width," & vbCrLf
sql = sql & "            SUM(CASE WHEN (TT.StockEffect = 1) AND (T.Transaction_Type = 3) THEN (TD.Quantity * TT.StockEffect) ELSE 0 END) AS openingValue," & vbCrLf
sql = sql & "            SUM(CASE WHEN (TT.StockEffect = 1) AND (T.Transaction_Type <> 3) THEN (TD.Quantity * TT.StockEffect) ELSE 0 END) AS inputvalue," & vbCrLf
sql = sql & "            SUM(CASE WHEN (TT.StockEffect = -1) THEN (TD.Quantity * TT.StockEffect) ELSE 0 END) AS outputValue" & vbCrLf
sql = sql & "        FROM dbo.Transaction_Details TD" & vbCrLf
sql = sql & "        INNER JOIN dbo.Transactions T ON TD.Transaction_ID = T.Transaction_ID" & vbCrLf
sql = sql & "        INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type" & vbCrLf
sql = sql & "        WHERE T.Transaction_Date >= @fromdate AND T.Transaction_Date <= @todate" & vbCrLf
sql = sql & "        AND (T.Storeid = @StoreId OR @StoreId IS NULL)" & vbCrLf
sql = sql & "        AND TT.StockEffect <> 0" & vbCrLf
'sql = sql & "        GROUP BY TD.Item_ID, ISNULL(TD.ColorID, 0), ISNULL(TD.ItemSize, ''), ISNULL(TD.ClassId, 0), ISNULL(TD.LotNO, ''), ISNULL(TD.UnitId, 0), ISNULL(TD.Length, 0), ISNULL(TD.Height, 0), ISNULL(TD.Width, 0)" & vbCrLf
sql = sql & "        GROUP BY TD.Item_ID, ISNULL(TD.ColorID, 0), ISNULL(TD.ItemSize, ''), ISNULL(TD.ClassId, 0), ISNULL(TD.LotNO, ''), ISNULL(TD.UnitId, 0)" & vbCrLf
sql = sql & "    )" & vbCrLf
sql = sql & "    INSERT @XTable" & vbCrLf
sql = sql & "    SELECT" & vbCrLf
sql = sql & "        c.Item_ID," & vbCrLf
sql = sql & "        c.ColorID," & vbCrLf
sql = sql & "        c.ItemSize," & vbCrLf
sql = sql & "        c.ClassId," & vbCrLf
sql = sql & "        c.LotNO," & vbCrLf
sql = sql & "        c.UnitId," & vbCrLf
sql = sql & "        TI.ItemCode," & vbCrLf
sql = sql & "        TI.ItemName," & vbCrLf
'sql = sql & "        c.Length," & vbCrLf
'sql = sql & "        c.Height," & vbCrLf
'sql = sql & "        c.Width," & vbCrLf
sql = sql & "        ISNULL(p.openingValue, 0) AS openingValue," & vbCrLf
sql = sql & "        ISNULL(p.inputvalue, 0) AS inputvalue," & vbCrLf
sql = sql & "        ISNULL(p.outputValue, 0) AS outputValue," & vbCrLf
sql = sql & "        ISNULL(o.openingBalance, 0) AS openingBalance" & vbCrLf
sql = sql & "    FROM AllComb c" & vbCrLf
sql = sql & "    LEFT JOIN TransInPeriod p ON c.Item_ID = p.Item_ID AND c.ColorID = p.ColorID AND c.ItemSize = p.ItemSize" & vbCrLf
'sql = sql & "    AND c.ClassId = p.ClassId AND c.LotNO = p.LotNO AND c.UnitId = p.UnitId AND c.Length = p.Length AND c.Height = p.Height AND c.Width = p.Width" & vbCrLf
sql = sql & "    AND c.ClassId = p.ClassId AND c.LotNO = p.LotNO AND c.UnitId = p.UnitId " & vbCrLf
sql = sql & "    LEFT JOIN Opening o ON c.Item_ID = o.Item_ID AND c.ColorID = o.ColorID AND c.ItemSize = o.ItemSize" & vbCrLf
'sql = sql & "    AND c.ClassId = o.ClassId AND c.LotNO = o.LotNO AND c.UnitId = o.UnitId AND c.Length = o.Length AND c.Height = o.Height AND c.Width = o.Width" & vbCrLf
sql = sql & "    AND c.ClassId = o.ClassId AND c.LotNO = o.LotNO AND c.UnitId = o.UnitId " & vbCrLf
sql = sql & "    INNER JOIN dbo.TblItems TI ON c.Item_ID = TI.ItemID" & vbCrLf
sql = sql & "    WHERE  (ISNULL(p.inputvalue, 0) + ISNULL(p.outputValue, 0) + ISNULL(o.openingBalance, 0)) <> 0  RETURN " & vbCrLf
sql = sql & "END"


db_createOrUpdateFuctionSQL "QryItemsInventry4", sql

'    sql = "    DROP FUNCTION QryItemsTransactionsTotals" & CHR(13)
'    Cn.Execute sql
'
'    sql = " CREATE FUNCTION QryItemsTransactionsTotals(@TransType int =0,@TransType2 int=0,@TransType3 int=0,@FromDate datetime ,@ToDate datetime ,@ItemID  as integer ,@Transaction_ID as float=null)" & CHR(13)
'    sql = sql & "RETURNS @xTable TABLE" & CHR(13)
'    sql = sql & "(" & CHR(13)
'    sql = sql & "ItemID int," & CHR(13)
'    sql = sql & "ItemCode nvarchar(50)," & CHR(13)
'    sql = sql & "ItemName nvarchar(4000)," & CHR(13)
'    sql = sql & "GroupID  int," & CHR(13)
'    sql = sql & "Total   money," & CHR(13)
'    sql = sql & "totalqty Float" & CHR(13)
'    sql = sql & ")" & CHR(13)
'    sql = sql & "AS" & CHR(13)
'    sql = sql & "Begin" & CHR(13)
'
'    sql = sql & "INSERT @xTable" & CHR(13)
'    sql = sql & "   Select ItemID,ItemCode,ItemName,GroupID,Sum(Total) as Totals,Sum(Quantity) as TotalQty" & CHR(13)
'    sql = sql & "from" & CHR(13)
'    sql = sql & "(" & CHR(13)
'    sql = sql & "SELECT TblItems.ItemID,TblItems.ItemCode, TblItems.ItemName,TblItems.GroupID," & CHR(13)
'    sql = sql & " Total= Transaction_Details.Quantity*Transaction_Details.Price " & CHR(13)
'      sql = sql & ",Transaction_Details.Quantity" & CHR(13)
'    sql = sql & "FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN" & CHR(13)
'    sql = sql & "dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
'    sql = sql & "WHERE (Transactions.Transaction_Type=@TransType  OR Transactions.Transaction_Type=@TransType2 OR Transactions.Transaction_Type=@TransType3" & CHR(13)
'    sql = sql & "or   Transactions.Transaction_Type=34  or     Transactions.Transaction_Type=15 )" & CHR(13)
'    sql = sql & "AND" & CHR(13)
'    sql = sql & "Transactions.Transaction_Date >=@FromDate" & CHR(13)
'    sql = sql & "AND" & CHR(13)
'    sql = sql & "Transactions.Transaction_Date <=@TODate" & CHR(13)
'      sql = sql & "AND" & CHR(13)
'   sql = sql & "  TblItems.ItemID =@ItemID" & CHR(13)
'   sql = sql & " and  Transactions.Transaction_ID<>isnull(@Transaction_ID,Transactions.Transaction_ID)" & CHR(13)
'
'  '  sql = sql & " and  Transaction_Details.UnitId = ISNULL(@UnitID,Transaction_Details.UnitId)" & CHR(13)
'   sql = sql & ")DrivTable" & CHR(13)
'
'
'    sql = sql & "Group By ItemID,ItemCode,ItemName,GroupID" & CHR(13)
'    sql = sql & "Return" & CHR(13)
'    sql = sql & " End" & CHR(13)


sql = ""
    sql = sql & "ALTER FUNCTION [dbo].[QryItemsTransactionsTotals]" & CHR(13)
    sql = sql & "(" & CHR(13)
    sql = sql & "    @TransType int = 0," & CHR(13)
    sql = sql & "    @TransType2 int = 0," & CHR(13)
    sql = sql & "    @TransType3 int = 0," & CHR(13)
    sql = sql & "    @FromDate datetime," & CHR(13)
    sql = sql & "    @ToDate datetime," & CHR(13)
    sql = sql & "    @ItemID int," & CHR(13)
    sql = sql & "    @Transaction_ID float = null" & CHR(13)
    sql = sql & ")" & CHR(13)
    sql = sql & "RETURNS @xTable TABLE" & CHR(13)
    sql = sql & "(" & CHR(13)
    sql = sql & "    ItemID int," & CHR(13)
    sql = sql & "    ItemCode nvarchar(50)," & CHR(13)
    sql = sql & "    ItemName nvarchar(4000)," & CHR(13)
    sql = sql & "    GroupID int," & CHR(13)
    sql = sql & "    Total money," & CHR(13)
    sql = sql & "    totalqty float" & CHR(13)
    sql = sql & ")" & CHR(13)
    sql = sql & "AS" & CHR(13)
    sql = sql & "BEGIN" & CHR(13)
    sql = sql & "    DECLARE @UseLastCount bit = 0;" & CHR(13)
    sql = sql & "    DECLARE @TreatUncountedAsZero bit = 0;" & CHR(13)
    sql = sql & "    SELECT TOP (1)" & CHR(13)
    sql = sql & "        @UseLastCount = CASE WHEN ISNULL(CostStartingGard,0)=1 THEN 1 ELSE 0 END," & CHR(13)
    sql = sql & "        @TreatUncountedAsZero = CASE WHEN ISNULL(TreatUncountedItemsAsZeroQty,0)=1 THEN 1 ELSE 0 END" & CHR(13)
    sql = sql & "    FROM dbo.TblOptions;" & CHR(13)
    sql = sql & "    ;WITH TargetCount AS (" & CHR(13)
    sql = sql & "        SELECT TargetCountDate = CAST(MAX(t.Transaction_Date) AS date)" & CHR(13)
    sql = sql & "        FROM dbo.Transactions t" & CHR(13)
    sql = sql & "        WHERE @UseLastCount = 1 AND t.Transaction_Type = 30 AND t.Transaction_Date < @FromDate" & CHR(13)
    sql = sql & "    ), Stores AS (" & CHR(13)
    sql = sql & "        SELECT s.StoreID FROM dbo.TblStore s" & CHR(13)
    sql = sql & "    ), LastCountPerStore AS (" & CHR(13)
    sql = sql & "        SELECT st.StoreID, x.LastCountTransId, x.LastCountDate" & CHR(13)
    sql = sql & "        FROM Stores st CROSS JOIN TargetCount tc" & CHR(13)
    sql = sql & "        OUTER APPLY (" & CHR(13)
    sql = sql & "            SELECT TOP (1) t.Transaction_ID AS LastCountTransId, t.Transaction_Date AS LastCountDate" & CHR(13)
    sql = sql & "            FROM dbo.Transactions t INNER JOIN dbo.Transaction_Details td ON td.Transaction_ID = t.Transaction_ID" & CHR(13)
    sql = sql & "            WHERE @UseLastCount = 1 AND t.StoreID = st.StoreID AND t.Transaction_Type = 30" & CHR(13)
    sql = sql & "              AND td.Item_ID = @ItemID AND tc.TargetCountDate IS NOT NULL" & CHR(13)
    sql = sql & "              AND t.Transaction_Date >= CAST(tc.TargetCountDate AS datetime)" & CHR(13)
    sql = sql & "              AND t.Transaction_Date < DATEADD(day,1,CAST(tc.TargetCountDate AS datetime))" & CHR(13)
    sql = sql & "            ORDER BY t.Transaction_Date DESC, t.Transaction_ID DESC" & CHR(13)
    sql = sql & "        ) x" & CHR(13)
    sql = sql & "    ), CountOpeningPerStore AS (" & CHR(13)
    sql = sql & "        SELECT lc.StoreID," & CHR(13)
    sql = sql & "            OpenQty = SUM(ISNULL(td.Quantity,0) / NULLIF(ISNULL(dbo.GetItemUnitFactor(td.Item_ID, td.UnitId),1),0))," & CHR(13)
    sql = sql & "            OpenVal = SUM(ISNULL(td.Quantity,0) * ISNULL(td.Price,0))" & CHR(13)
    sql = sql & "        FROM LastCountPerStore lc INNER JOIN dbo.Transaction_Details td ON td.Transaction_ID = lc.LastCountTransId" & CHR(13)
    sql = sql & "        AND td.Item_ID = @ItemID WHERE lc.LastCountTransId IS NOT NULL GROUP BY lc.StoreID" & CHR(13)
    sql = sql & "    ), StoreStart AS (" & CHR(13)
    sql = sql & "        SELECT st.StoreID, lc.LastCountDate," & CHR(13)
    sql = sql & "            OpeningSource = CASE WHEN @UseLastCount = 1 AND lc.LastCountTransId IS NOT NULL THEN 'COUNT'" & CHR(13)
    sql = sql & "                                 WHEN @UseLastCount = 1 AND lc.LastCountTransId IS NULL AND @TreatUncountedAsZero = 1 THEN 'ZERO'" & CHR(13)
    sql = sql & "                                 ELSE 'HISTORICAL' END," & CHR(13)
    sql = sql & "            StartDate = CASE WHEN @UseLastCount = 1 AND lc.LastCountTransId IS NOT NULL THEN DATEADD(day,1,CAST(CAST(lc.LastCountDate AS date) AS datetime))" & CHR(13)
    sql = sql & "                             WHEN @UseLastCount = 1 AND lc.LastCountTransId IS NULL AND @TreatUncountedAsZero = 1 THEN @FromDate" & CHR(13)
    sql = sql & "                             ELSE CAST('19000101' AS datetime) END" & CHR(13)
    sql = sql & "        FROM Stores st LEFT JOIN LastCountPerStore lc ON lc.StoreID = st.StoreID" & CHR(13)
    sql = sql & "    ), Movements AS (" & CHR(13)
    sql = sql & "        SELECT t.StoreID, t.Transaction_Date, tt.StockEffect," & CHR(13)
    sql = sql & "            UnitFactor = ISNULL(dbo.GetItemUnitFactor(td.Item_ID, td.UnitId),1)," & CHR(13)
    sql = sql & "            QtyBase = ISNULL(td.Quantity,0)," & CHR(13)
    sql = sql & "            ValSigned = ROUND(ISNULL(td.Quantity,0) * ISNULL(td.Price,0) * ISNULL(tt.StockEffect,0), 2)" & CHR(13)
    sql = sql & "        FROM dbo.Transactions t INNER JOIN dbo.TransactionTypes tt ON tt.Transaction_Type = t.Transaction_Type" & CHR(13)
    sql = sql & "        INNER JOIN dbo.Transaction_Details td ON td.Transaction_ID = t.Transaction_ID" & CHR(13)
    sql = sql & "        WHERE td.Item_ID = @ItemID AND ISNULL(tt.StockEffect,0) <> 0 AND t.Transaction_Type <> 30" & CHR(13)
    sql = sql & "          AND t.Transaction_ID <> ISNULL(@Transaction_ID, t.Transaction_ID)" & CHR(13)
    sql = sql & "          AND (t.Transaction_Type IN (@TransType, @TransType2, @TransType3, 34, 15))" & CHR(13)
    sql = sql & "    ), Classified AS (" & CHR(13)
    sql = sql & "        SELECT ss.StoreID," & CHR(13)
    sql = sql & "            OpenQtySigned = CASE WHEN m.Transaction_Date >= ss.StartDate AND m.Transaction_Date < @FromDate THEN (m.QtyBase / NULLIF(ISNULL(m.UnitFactor,1),0)) * m.StockEffect ELSE 0 END," & CHR(13)
    sql = sql & "            OpenValSigned = CASE WHEN m.Transaction_Date >= ss.StartDate AND m.Transaction_Date < @FromDate THEN m.ValSigned ELSE 0 END," & CHR(13)
    sql = sql & "            PeriodQtySigned = CASE WHEN m.Transaction_Date >= @FromDate AND m.Transaction_Date <= @ToDate THEN (m.QtyBase / NULLIF(ISNULL(m.UnitFactor,1),0)) * m.StockEffect ELSE 0 END," & CHR(13)
    sql = sql & "            PeriodValSigned = CASE WHEN m.Transaction_Date >= @FromDate AND m.Transaction_Date <= @ToDate THEN m.ValSigned ELSE 0 END" & CHR(13)
    sql = sql & "        FROM StoreStart ss LEFT JOIN Movements m ON m.StoreID = ss.StoreID" & CHR(13)
    sql = sql & "    ), AggStores AS (" & CHR(13)
    sql = sql & "        SELECT TotalQtyAll = SUM(OpenQtySigned) + SUM(PeriodQtySigned) + SUM(ISNULL(co.OpenQty,0))," & CHR(13)
    sql = sql & "               TotalValAll = SUM(OpenValSigned) + SUM(PeriodValSigned) + SUM(ISNULL(co.OpenVal,0))" & CHR(13)
    sql = sql & "        FROM Classified c LEFT JOIN CountOpeningPerStore co ON co.StoreID = c.StoreID" & CHR(13)
    sql = sql & "    )" & CHR(13)
    sql = sql & "    INSERT @xTable" & CHR(13)
    sql = sql & "    SELECT i.ItemID, i.ItemCode, i.ItemName, i.GroupID, SUM(a.TotalValAll), SUM(a.TotalQtyAll)" & CHR(13)
    sql = sql & "    FROM dbo.TblItems i CROSS JOIN AggStores a WHERE i.ItemID = @ItemID" & CHR(13)
    sql = sql & "    Group By ItemID,ItemCode,ItemName,GroupID" & CHR(13)
    sql = sql & "    Return" & CHR(13)
    sql = sql & " End" & CHR(13)
    db_createOrUpdateFuctionSQL "QryItemsTransactionsTotals", sql


   sql = "    DROP FUNCTION GetItemqtytodate2015" & CHR(13)
    Cn.Execute sql
sql = ""
sql = sql & " Create FUNCTION GetItemqtytodate2015" & CHR(13)
sql = sql & " (" & CHR(13)
sql = sql & "     @Todate       DATETIME," & CHR(13)
sql = sql & "     @itemid       AS integer," & CHR(13)
sql = sql & "     @transid1     AS FLOAT," & CHR(13)
sql = sql & "     @transid2     AS FLOAT" & CHR(13)
sql = sql & " )" & CHR(13)
sql = sql & " RETURNS Float" & CHR(13)
sql = sql & " AS" & CHR(13)

sql = sql & " Begin" & CHR(13)
sql = sql & "     RETURN (" & CHR(13)
sql = sql & "         SELECT SUM(" & CHR(13)
sql = sql & "                    dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect" & CHR(13)
sql = sql & "                ) AS SumQty" & CHR(13)
sql = sql & "         From dbo.Transaction_Details" & CHR(13)
sql = sql & "                INNER JOIN dbo.Transactions" & CHR(13)
sql = sql & "                     ON  dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
sql = sql & "                INNER JOIN dbo.TransactionTypes" & CHR(13)
sql = sql & "                     ON  dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type" & CHR(13)
sql = sql & "         Where (dbo.TransactionTypes.StockEffect <> 0)" & CHR(13)
sql = sql & "                AND (dbo.Transaction_Details.Item_ID = @itemid)" & CHR(13)
sql = sql & "                AND (dbo.Transactions.Transaction_Date < @Todate)" & CHR(13)
sql = sql & "                AND (" & CHR(13)
sql = sql & "                        dbo.Transaction_Details.Transaction_ID <> @transid1" & CHR(13)
sql = sql & "                        AND dbo.Transaction_Details.Transaction_ID <> @transid2" & CHR(13)
sql = sql & "                    )" & CHR(13)
sql = sql & "     )" & CHR(13)

sql = sql & "                       End" & CHR(13)


    db_createOrUpdateFuctionSQL "GetItemqtytodate2015", sql
    
    

    
       sql = "    DROP FUNCTION GetItemCostPrice" & CHR(13)
    Cn.Execute sql
sql = ""
sql = sql & " Create FUNCTION GetItemCostPrice" & CHR(13)
sql = sql & " (" & CHR(13)
sql = sql & "     @fromdate       DATETIME," & CHR(13)
sql = sql & "     @Todate       DATETIME," & CHR(13)
sql = sql & "     @itemid       AS integer " & CHR(13)
sql = sql & " ) " & CHR(13)
sql = sql & " RETURNS Float" & CHR(13)
sql = sql & " AS" & CHR(13)

sql = sql & " Begin" & CHR(13)
sql = sql & "     RETURN (" & CHR(13)
sql = sql & "         SELECT ROUND(Total / totalqty, 5) AS AvCost " & CHR(13)
sql = sql & "         FROM dbo.QryItemsTransactionsTotals(28, 3, 20, @fromdate, @Todate, @itemid,0) " & CHR(13)

sql = sql & "                WHERE ItemID = @itemid " & CHR(13)
sql = sql & "     )" & CHR(13)

sql = sql & "                       End" & CHR(13)


    db_createOrUpdateFuctionSQL "GetItemCostPrice", sql
End Function

Function UpdateDataBasePart29()

    If DB_CreateTable("TblContractInstallDisco2", True, "id ", True) = True Then
        DB_CreateField "TblContractInstallDisco2", "valuewithout", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblContractInstallDisco2", "VatPerc", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblContractInstallDisco2", "VatValue", adDouble, adColNullable, , , "    ", False, True
    
        DB_CreateField "TblContractInstallDisco2", "ValueAfterDiscount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblContractInstallDisco2", "Discount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblContractInstallDisco2", "DiscountValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblContractInstallDisco2", "Value", adDouble, adColNullable, , , "    ", False, True
        'DB_CreateField "TblContractInstallDisco2", "[Select]", adBoolean, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "Select", adBoolean, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "PaymentNo", adInteger, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "MasterNo", adInteger, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "MasterID", adInteger, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "Cont", adInteger, adColNullable, , , "                ", False, True
        DB_CreateField "TblContractInstallDisco2", "Ser", adInteger, adColNullable, , , "                ", False, True
    
        DB_CreateField "TblContractInstallDisco2", "RecDateH", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblContractInstallDisco2", "AllowDateH", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
        DB_CreateField "TblContractInstallDisco2", "DMY", adVarWChar, adColNullable, 255, , "„·«ÕŸ«       ", False, True, , True
    
        DB_CreateField "TblContractInstallDisco2", "AllowDate", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "TblContractInstallDisco2", "RecDate", adDBTimeStamp, adColNullable, , , "", False, True

    End If

    'Cn.Execute "Delete TBLTYPEIMAGE"
    'Cn.Execute "Delete TBLTYPEIMAGE2"
    Dim s       As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select * from TblTypeImage "
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic

    If rsDummy.EOF Then
        s = " INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "     VALUES(1, '«· √”Ì”', NULL);" & vbNewLine & vbNewLine
    
        s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "         VALUES(2, '«·÷—«∆»', NULL);" & vbNewLine
        'GO
    
        s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "         VALUES(3, '«·„—«Ã⁄Â', NULL);" & vbNewLine
        'GO
    
        s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "         VALUES(4, '„·›«  œ—«”… «·ÃœÊÏ', NULL);" & vbNewLine
        ' GO
    
        s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "         VALUES(5, '«·‘∆Ê‰ «·ﬁ«‰Ê‰Ì…', NULL);" & vbNewLine
        ' GO
    
        s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        s = s & "         VALUES(6, '«· «„Ì‰«  «·«Ã „«⁄Ì…', NULL);" & vbNewLine
        ' GO
        '
        '    s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        '    s = s & "         VALUES(7, '÷—«∆» ﬁÌ„… „÷«›… ', NULL);" & vbNewLine
        '    ' GO
        '
        '    s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        '    s = s & "         VALUES(8, '÷—Ì»… «·œ„€…', NULL);" & vbNewLine
        '    'GO
        '
        '    s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        '    s = s & "         VALUES(9, '÷—Ì»… ﬂ”» «·⁄„· ', NULL);" & vbNewLine
        '    'GO
        '
        '    s = s & "     INSERT INTO [TBLTYPEIMAGE]([ID], [Name], [NameE])" & vbNewLine
        '    s = s & "         VALUES(10, '«·÷—Ì»… «·⁄ﬁ«—Ì…', NULL);" & vbNewLine
        'GO

        Cn.Execute s

        s = " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(1, '⁄ﬁœ «·‘—ﬂ…  ', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(2, '’ÕÌ›… «·«” À„«—', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(3, '⁄ﬁÊœ  ⁄œÌ·', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(4, '«·»ÿ«ﬁ… «·÷—Ì»Ì… ', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(5, '«·”Ã· «· Ã«—Ì', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(6, '«·»ÿ«ﬁ… «·«” Ì—«œÌ…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(7, '«·»ÿ«ﬁ… «· ’œÌ—Ì…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(8, '⁄ﬁœ «·«ÌÃ«—', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(9, '⁄ﬁœ „·ﬂÌ…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(10, '«·Ã„⁄Ì… «·⁄„Ê„Ì… «·⁄«œÌ…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(11, '«·Ã„⁄Ì«  «·⁄„Ê„Ì… €Ì— «·⁄«œÌ…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(12, '„Õ«÷— „Ã«·” «·«œ«—…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(13, 'Õ«›Ÿ… «·»—Ìœ', NULL, 1);"
        'GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(14, '„—«”·«  ÂÌ∆… «·«” À„«—', NULL, 1);"
        'GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(15, '„—«›ﬁ«  ÊŒÿ«»«  ÂÌ∆… «·—ﬁ«»… «·„«·Ì…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(16, '„’— «·„ﬁ«’…', NULL, 1);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(17, '‰„«–Ã «·÷—«∆»', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(18, '«·ÿ⁄‰ «·÷—Ì»Ì', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(19, 'ﬁ—«—«  ·Ã«‰ «·ÿ⁄‰', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(20, '«Œÿ«—«  «·÷—«∆»', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(21, '„–ﬂ—… «·›Õ’ «·÷—Ì»Ì', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(22, '„Õ«÷— «·«⁄„«·', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(23, '«Ì’«·«  ”œ«œ ÷—Ì»Ì', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(24, '«·«ﬁ—«— «·÷—Ì»Ì «·”‰ÊÌ', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(25, ' ﬁ—Ì— «·„—«Ã⁄Â «·—»⁄ ”‰ÊÌ', '', 3);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(26, '„—›ﬁ«  „·›', '', 3);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(27, 'œ—«”… «·ÃœÊÌ', NULL, 4);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(28, '⁄—Ì÷… «·œ⁄ÊÌ', '', 5);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(29, '‘Â«œ… „‰ «·ÃœÊ·', '', 5);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(30, '„—›ﬁ«  «·„·›', '', 5);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(31, '«” „«—… 2  √„Ì‰«  «Ã „«⁄Ì…', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(32, '«” „«—… 1  √„Ì‰«  ··⁄«„·Ì‰', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(33, '«” „«—… 6  «„Ì‰« ', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(34, '«” ﬁ«·…', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(35, '«” ·«„ „” Õﬁ« ', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(36, '«Ì’«·«  ”œ«œ «· «„Ì‰« ', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(37, '«ﬁ—«—«  «·÷—Ì»… «·‘Â—Ì…', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(38, '„–ﬂ—… «·›Õ’ «·÷—Ì»Ì', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(39, '„Õ«÷— «⁄„«· ›Õ’ ÷—Ì»Ì', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(40, '„—«”·«  ÷—Ì»… «·ﬁÌ„… «·„÷«›…', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(41, '‘Â«œ…  ”ÃÌ· «·ﬁÌ„… «·„÷«›…', '', 6);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(42, '„–ﬂ—… «·›Õ’ - œ„€…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(43, '«·‰„«–Ã «·÷—Ì»Ì… - œ„€…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(44, '«·ÿ⁄Ê‰ «·÷—Ì»Ì… - œ„€…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(45, ' ”ÊÌ… ﬂ”» «·⁄„·', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(46, '«ﬁ—«— ÷—Ì»Ì ‰„Ê–Ã (4)', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(47, '‰„«–Ã („ÿ«·»…) ﬂ”» ⁄„·', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(48, '«·ÿ⁄‰ «·÷—Ì»Ì ⁄‰ „ÿ«·»… ﬂ”» «·⁄„·', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(49, '„ÿ«·»… «·÷—Ì»… «·⁄ﬁ«—Ì…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(50, '«·ÿ⁄‰ «·÷—Ì»Ì ', '', 2);"

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine & vbNewLine
        s = s & "     VALUES(51, '„—«”·«  ÷—Ì»Ì…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(52, '„–ﬂ—… «·›Õ’ - œ„€…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(53, '«·‰„«–Ã «·÷—Ì»Ì… - œ„€…', '', 2);"
        ' GO

        s = s & " INSERT INTO [TBLTYPEIMAGE2]([ID], [Name], [NameE], [MasterId])" & vbNewLine
        s = s & "     VALUES(54, '«·ÿ⁄Ê‰ «·÷—Ì»Ì… - œ„€…', '', 2);"
        ' GO

        Cn.Execute s

    End If

End Function

Function UpdateDataBasePart30()
    On Error Resume Next
    Dim New_View As String
    Dim s        As String
   
    
    '*************Ship *************
      If DB_CreateTable("TBLLCHistory", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "TBLLCHistory", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCHistory", "TblLCID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCHistory", "serial", adInteger, adColNullable, , , ""
        
        DB_CreateField "TBLLCHistory", "Code", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "TBLLCHistory", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "TBLLCHistory", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True
        
        DB_CreateField "TBLLCHistory", "GuaranteeAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLLCHistory", "AmountPlus", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLLCHistory", "AmountMin", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLLCHistory", "Total", adDouble, adColNullable, , , "    ", False, True
        
        
    
    
    
    End If
    DB_CreateField "TblOptions", "ShowPrinterDialoge", adBoolean, adColNullable, , , "", False, True
     DB_CreateField "TblCustemers", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "DOUBLE_ENTREY_VOUCHERS", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "notes_all", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Transactions", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    DB_CreateField "Notes", "IsHiddenInv", adBoolean, adColNullable, , , "        ", False, True
    
   
   DB_CreateField "TblEmpAllocations", "ActivityTypeId", adInteger, adColNullable, , , ""
   
    UpdateAccountProc
    UpdateCostriceProcedure
    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 5002 ,'›« Ê—… ⁄ﬁÊœ «·«„·«ﬂ  ' ,'    Invoice  ' ", "NotesType", 5002
    add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 5003 ,'›« Ê—… ÷—Ì»Ì…  ' ,'    Invoice  ' ", "NotesType", 5003
    DB_CreateField "TblOptions", "BigUserPw2", adVarWChar, adColNullable, 400, , "", False, True, , True
   DB_CreateField "tmpPos33", "NetValue5", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue6", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue7", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue8", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue9", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue10", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "tmpPos33", "NetValue4", adDouble, adColNullable, , , "    ", False, True
   
DB_CreateField "tmpPos33", "ID4", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID5", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID6", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID7", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID8", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID9", adInteger, adColNullable, , , ""
DB_CreateField "tmpPos33", "ID10", adInteger, adColNullable, , , ""
    
    UpdateEmpVoCation2
    CreateOrUpdateTrigger
    
'UpdateCostriceProcedureByStores
DB_CreateField "tblContractInsAllocationsDetails", "OldNoteSerial1H", adVarWChar, adColNullable, 250, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "NoteSerial1H", adVarWChar, adColNullable, 250, , "", False, True, , True
DB_CreateField "TblContractInstallments", "NoteSerial1H", adVarWChar, adColNullable, 250, , "", False, True, , True


DB_CreateField "TblEmpDepartments", "a7", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a7", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a29", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a30", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a74", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a93", adVarWChar, adColNullable, 250, , "", False, True, , True
 DB_CreateField "TblEmpDepartments", "a65", adVarWChar, adColNullable, 250, , "", False, True, , True
 
 
DB_CreateField "Transactions", "ExportDate", adDBTimeStamp, adColNullable, , , ""
DB_CreateField "Transactions", "ExportingCompany", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "BillOfLadingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "InvoiceNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "DeliveryPermitNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "OrderNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
 
 
 
DB_CreateField "Transactions", "CountryOfDest", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "CountryOfOrigin", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "Currency", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "ExportingCompany", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "ExportingCompany2", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "CustomsBrokerOrAuthorized", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "ImporterID", adInteger, adColNullable, , , ""

DB_CreateField "Transactions", "HarborID", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "HarborID2", adInteger, adColNullable, , , ""
 
 DB_CreateField "Transactions", "CommitmentDetails", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "Weight", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "NumberOfPackages", adInteger, adColNullable, , , "", False, True
DB_CreateField "Transactions", "OrderCompletionDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "FeeAmount", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CommitmentNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "CommitmentDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "ImporterNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "CustomsBrokerNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "CustomsBrokerOrAuthorized", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "ImportDeclarationNumber", adVarWChar, adColNullable, 400, , "", False, True, , True

DB_CreateField "Transactions", "CarrierOrShippingAgent", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "ExportPort", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "LoadingPort", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "ExportCountryCode", adVarWChar, adColNullable, 400, , "", False, True, , True




DB_CreateField "Transactions", "TotalInvoice", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CustomsReceipt", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "ExportDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "BillOfLadingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "InvoiceNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "DeliveryPermitNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "OrderNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "Transactions", "TransactionStatusDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "ExpectedArrivalDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "ActualArrivalDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "DocumentsSentDate", adDBTimeStamp, adColNullable, , , "", False, True

 

DB_CreateField "project_billl", "PerformanceBond", adDouble, adColNullable, , , "    ", False, True
 
    If DB_CreateTable("tblEInvoice", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        'DB_CreateField "tblEInvoice", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tblEInvoice", "InvoiceID", adInteger, adColNullable, , , ""
        DB_CreateField "tblEInvoice", "DefaultInvoicetype", adInteger, adColNullable, , , ""
      
       DB_CreateField "tblEInvoice", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "", False, True, , True

        DB_CreateField "tblEInvoice", "IssueDate", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tblEInvoice", "IssueTim", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tblEInvoice", "DocumentCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "TaxCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "StreetName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "BuildingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "CityName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "PostalZone", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "CitySubdivisionName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "RegistrationName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "CompanyID", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "ItemName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice", "Qty", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice", "Price", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice", "CoCRCode", adVarWChar, adColNullable, 400, , "", False, True, , True
       
        DB_CreateField "tblEInvoice", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice", "VatValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "tblEInvoice", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
        DB_CreateField "tblEInvoice", "serial", adInteger, adColNullable, , , ""
                    DB_CreateField "tblEInvoice", "ExcelRow", adInteger, adColNullable, , , ""
    DB_CreateField "tblEInvoice", "ExcelFile", adVarWChar, adColNullable, 400, , "", False, True, , True

        
 DB_CreateField "tblEInvoice", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tblEInvoice", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tblEInvoice", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
   
        DB_CreateField "tblEInvoice", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "tblEInvoice", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "tblEInvoice", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblEInvoice", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "tblEInvoice", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

    
    
    
    End If
    

'DB_CreateField "tblEInvoice", "InvoiceID", adInteger, adColNullable, , , ""

DB_CreateField "tblEInvoice", "IsFromUser", adInteger, adColNullable, , , , False, True
    
    DB_CreateField "tblEInvoice", "TaxCategoryPercent", adDouble, adColNullable, , , ""
    DB_CreateField "tblEInvoice", "zatcaStatus", adInteger, adColNullable, , , , False, True
    DB_CreateField "tblEInvoice", "Transaction_ID", adInteger, adColNullable, , , ""
   DB_CreateField "tblEInvoice", "InvoiceID", adInteger, adColNullable, , , ""
      DB_CreateField "tblEInvoice2", "warrningmessage", adVarWChar, adColNullable, 4000, , "", False, True, , True
   DB_CreateField "tblEInvoice", "warrningmessage", adVarWChar, adColNullable, 4000, , "", False, True, , True




DB_CreateField "tblEInvoice", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "tblEInvoice", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True

DB_CreateField "tblEInvoice", "AdditionalStreetName", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice", "PlotIdentification", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice", "CountrySubentity", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice", "IdentificationCode", adVarWChar, adColNullable, 4000, , "", False, True, , True

 DB_CreateField "tblEInvoice", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True

  
    
     
    If DB_CreateTable("tblEInvoice2", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        'DB_CreateField "tblEInvoice2", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tblEInvoice2", "InvoiceID", adInteger, adColNullable, , , ""
        DB_CreateField "tblEInvoice2", "DefaultInvoicetype", adInteger, adColNullable, , , ""
      
       DB_CreateField "tblEInvoice2", "zatcaStatus", adInteger, adColNullable, , , , False, True
        DB_CreateField "tblEInvoice2", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "IssueDate", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tblEInvoice2", "IssueTim", adDBTimeStamp, adColNullable, , , "", False, True
        DB_CreateField "tblEInvoice2", "DocumentCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "TaxCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "StreetName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "BuildingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "CityName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "PostalZone", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "CitySubdivisionName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "RegistrationName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "CompanyID", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "ItemName", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "Qty", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice2", "Price", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice2", "CoCRCode", adVarWChar, adColNullable, 400, , "", False, True, , True
       
        DB_CreateField "tblEInvoice2", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice2", "VatValue", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblEInvoice2", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "tblEInvoice2", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
        DB_CreateField "tblEInvoice2", "serial", adInteger, adColNullable, , , ""
        
        
 DB_CreateField "tblEInvoice2", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tblEInvoice2", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tblEInvoice2", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
   
        
DB_CreateField "tblEInvoice2", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "tblEInvoice2", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice2", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice2", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice2", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice2", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice2", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice2", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblEInvoice2", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "tblEInvoice2", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

    
    
    
    End If
    DB_CreateField "tblEInvoice2", "ExcelFile", adVarWChar, adColNullable, 400, , "", False, True, , True
    DB_CreateField "tblEInvoice2", "ExcelRow", adInteger, adColNullable, , , ""
    DB_CreateField "tblEInvoice2", "Transaction_ID", adInteger, adColNullable, , , ""
   DB_CreateField "tblEInvoice2", "InvoiceID", adInteger, adColNullable, , , ""
     
   DB_CreateField "tblEInvoice2", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "tblEInvoice2", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblEInvoice2", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblEInvoice2", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice2", "warrningmessage", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice2", "AdditionalStreetName", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice2", "PlotIdentification", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice2", "CountrySubentity", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblEInvoice2", "IdentificationCode", adVarWChar, adColNullable, 4000, , "", False, True, , True

 DB_CreateField "tblEInvoice2", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True

   
   DB_CreateField "tblEInvoice2", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "", False, True, , True
   
   
   DB_CreateField "tblEInvoice", "Identificationid", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "tblEInvoice2", "Identificationid", adVarWChar, adColNullable, 255, , "", False, True, , True
   
   DB_CreateField "tblEInvoice", "schemeID", adVarWChar, adColNullable, 255, , "", False, True, , True
   DB_CreateField "tblEInvoice2", "schemeID", adVarWChar, adColNullable, 255, , "", False, True, , True
   

   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
    DB_CreateField "tblEInvoice", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    
    
    DB_CreateField "project_billl", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True
    
 
  
   
DB_CreateField "TblOptions", "CustVatNoMandatory", adBoolean, adColNullable, , , "", False, True

DB_CreateField "TblVocationEntitlements", "DaysCountPay", adDouble, adColNullable, , , "      ", False, True

DB_CreateField "TblEmployee", "balanceH1", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "balanceH2", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "balanceH3", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "TblEmployee", "balanceH4", adDouble, adColNullable, , , "      ", False, True

DB_CreateField "TblVocationEntitlements", "ch9", adBoolean, adColNullable, , , "", False, True

DB_CreateField "notes_all", "ContainerNo", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TblTravDueK", "ContainerNo", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TblTravDueKDet", "ContainerNo", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TblOrderUpload", "ContainerNo", adVarWChar, adColNullable, 400, , "", False, True, , True

    
         If DB_CreateTable("TBLdAILYSALES", True, "ID", True) = True Then
            DB_CreateField "TBLdAILYSALES", "transName", adDouble, adColNullable, , , , False, True
            DB_CreateField "TBLdAILYSALES", "todaytotals", adDouble, adColNullable, , , , False, True
            DB_CreateField "TBLdAILYSALES", "minthtotal", adDouble, adColNullable, , , , False, True
            DB_CreateField "TBLdAILYSALES", "YearsTotal", adDouble, adColNullable, , , , False, True
            DB_CreateField "TBLdAILYSALES", "Comparison", adDouble, adColNullable, , , , False, True
            
       End If

DB_CreateField "TblCorBalaCusDet", "balancetype", adInteger, adColNullable, , , ""

 
       
DB_CreateField "TblCustemers", "StreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
              
              DB_CreateField "TblCustemers", "StreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "AdditionalStreetName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "BuildingNumber", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "PlotIdentification", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "CityName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "PostalZone", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "CountrySubentity", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "CitySubdivisionName", adVarWChar, adColNullable, 255, , "", False, True, , True
       DB_CreateField "TblCustemers", "IdentificationCode", adVarWChar, adColNullable, 255, , "", False, True, , True
       
       DB_CreateField "TblCustemers", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
       
DB_CreateField "TblDoCumentsTypes", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True


DB_CreateField "transactions", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "Transactions", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Transactions", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Transactions", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Transactions", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Transactions", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "Transactions", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

    




DB_CreateField "Transactions", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Transactions", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True

DB_CreateField "TblOptions", "SalesBoxIDReturn", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "allowancechargeAmount", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "allowancechargeAllowanceChargeReason", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "Transactions", "allowancechargeTaxCategoryid", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "allowancechargeTaxCategoryPercent", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "LegalMonetaryTotalPayableAmount", adDouble, adColNullable, , , , False, True
DB_CreateField "Transactions", "LegalMonetaryTotalPrepaidAmount", adDouble, adColNullable, , , , False, True
 
 
         If DB_CreateTable("transactionsVatDetails", True, "ID", True) = True Then
         
        
 DB_CreateField "transactionsVatDetails", "SingedXML", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
 DB_CreateField "transactionsVatDetails", "EncodedInvoice", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
  DB_CreateField "transactionsVatDetails", "InvoiceHash", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
   DB_CreateField "transactionsVatDetails", "UUID", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
    DB_CreateField "transactionsVatDetails", "QRCode", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
    DB_CreateField "transactionsVatDetails", "PIH", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
    DB_CreateField "transactionsVatDetails", "SingedXMLFileName", adLongVarWChar, adColNullable, 10000, , "", False, True, , True
    
 End If
 
 DB_CreateField "transactionsVatDetails", "uuidCounter", adDouble, adColNullable, , , "      ", False, True
  DB_CreateField "transactionsVatDetails", "QrCodeData", adVarWChar, adColNullable, 255, , "      ", False
DB_CreateField "transactionsVatDetails", "QrCodeDataPath", adVarWChar, adColNullable, 255, , "      ", False
DB_CreateField "transactionsVatDetails", "QrCodeImage", adLongVarBinary, adColNullable, , , "      ", False, True
 
 
DB_CreateField "transactionsVatDetails", "Doctype", adInteger, adColNullable, , , , False, True

 DB_CreateField "Transaction_Details", "SelectedPOnumber", adVarWChar, adColNullable, 20, , "      ", False, True, , True
 DB_CreateField "transactionsVatDetails", "IsDeleted", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Transactions", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_updateField "transactionsVatDetails", "QrCodeData", "nvarchar(4000)   "



DB_CreateField "Transactions", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True


DB_CreateField "project_billl", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "project_billl", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "project_billl", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "project_billl", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True



DB_CreateField "project_billl", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "project_billl", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "project_billl", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "project_billl", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "project_billl", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "project_billl", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_billl", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "project_billl", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "project_billl", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "project_billl", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "project_billl", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_billl", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_billl", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_billl", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "project_billl", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "project_billl", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "project_billl", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "project_billl", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "project_billl", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "project_billl", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True


DB_CreateField "project_billl", "Invoicetype", adInteger, adColNullable, , , "    ", False, True




DB_CreateField "Transactions", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True








DB_CreateField "Transactions", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True












   DB_CreateField "TblHandWages", "GeneralTotal", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblHandWages", "TotalDisc", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblHandWages", "TotalBVat", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblHandWages", "TotalVat", adDouble, adColNullable, , , "    ", False, True
   DB_CreateField "TblHandWages", "TotalNet", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblHandWages", "zatcaStatus", adInteger, adColNullable, , , , False, True


DB_CreateField "TblHandWages", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "TblHandWages", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblHandWages", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "TblHandWages", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

DB_CreateField "TblHandWages", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblHandWages", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblHandWages", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "TblHandWages", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "TblHandWages", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "TblHandWages", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblHandWages", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblHandWages", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblHandWages", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblHandWages", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblHandWages", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "TblHandWages", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "TblHandWages", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "TblHandWages", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "TblHandWages", "CusID", adInteger, adColNullable, , , , False, True

DB_CreateField "TblHandWages", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "TblHandWages", "Invoicetype", adInteger, adColNullable, , , "    ", False, True




DB_CreateField "TblHandWages", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True





    DB_CreateField "tblContractInsAllocationsDetails", "NoteSerial1H", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "tblContractInsAllocationsDetails", "zatcaStatus", adInteger, adColNullable, , , , False, True


DB_CreateField "tblContractInsAllocationsDetails", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "tblContractInsAllocationsDetails", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

DB_CreateField "tblContractInsAllocationsDetails", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "CusID", adInteger, adColNullable, , , , False, True
DB_CreateField "tblContractInsAllocationsDetails", "HNoteSerial1", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "tblContractInsAllocationsDetails", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "tblContractInsAllocationsDetails", "Invoicetype", adInteger, adColNullable, , , "    ", False, True




DB_CreateField "tblContractInsAllocationsDetails", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True




DB_CreateField "transactionsVatDetails", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "Notes", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "Notes", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "Notes", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "Notes", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "Notes", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "Notes", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

DB_CreateField "Notes", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "Notes", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "Notes", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "Notes", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "Notes", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Notes", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Notes", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "Notes", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Notes", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Notes", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Notes", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Notes", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Notes", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Notes", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Notes", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "Notes", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "Notes", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "Notes", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "Notes", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True



DB_CreateField "Notes", "Invoicetype", adInteger, adColNullable, , , "    ", False, True






DB_CreateField "notes_all", "TableName", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "notes_all", "last_changed", adDBTimeStamp, adColNullable, , , "", False, True

DB_CreateField "notes_all", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True

DB_CreateField "notes_all", "warrningmessage", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "notes_all", "ErrorMessageS", adVarWChar, adColNullable, 4000, , "„·«ÕŸ«       ", False, True, , True
DB_CreateField "notes_all", "RecTime", adDBTimeStamp, adColNullable, , , "   ", False, True

DB_CreateField "notes_all", "Doctype", adInteger, adColNullable, , , , False, True
DB_CreateField "notes_all", "Currency_id", adInteger, adColNullable, , , , False, True
DB_CreateField "notes_all", "Currency_rate", adDouble, adColNullable, , , , False, True
DB_CreateField "notes_all", "zatcaStatus", adInteger, adColNullable, , , , False, True
DB_CreateField "notes_all", "DateRec", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "notes_all", "CIBAN", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "notes_all", "InvoiceTypeCodeID", adDouble, adColNullable, , , , False, True
DB_CreateField "notes_all", "InvoiceTypeCodename", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "notes_all", "DocumentCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "notes_all", "TaxCurrencyCode", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "notes_all", "AdditionalDocumentReferencePIH", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "notes_all", "InvoiceDocumentReferenceID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "notes_all", "AdditionalDocumentReferenceICVUUID", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "notes_all", "ActualDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "notes_all", "LatestDeliveryDate", adDBTimeStamp, adColNullable, , , "", False, True
DB_CreateField "notes_all", "PaymentMeansCode", adDouble, adColNullable, , , , False, True
DB_CreateField "notes_all", "InstructionNote", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
DB_CreateField "notes_all", "Iban", adVarWChar, adColNullable, 255, , "", False, True, , True
DB_CreateField "notes_all", "paymentnote", adVarWChar, adColNullable, 4000, , "", False, True, , True



DB_CreateField "notes_all", "Invoicetype", adInteger, adColNullable, , , "    ", False, True



 DB_CreateField "TblCustemers", "export", adInteger, adColNullable, , , "    ", False, True
 
 
    DB_CreateField "tblOPtions", "AllowOpticalCycle", adBoolean, adColNullable, , , "", False, True
 
 
DB_updateField "transactionsVatDetails", "QrCodeData", "nvarchar(4000)   "
DB_CreateField "transactionsVatDetails", "IsDeleted", adInteger, adColNullable, , , "    ", False, True
 
    DB_CreateField "transactionsVatDetails", "IsDeleted", adInteger, adColNullable, , , "    ", False, True

DB_CreateField "TblBranchesData", "CRN", adVarWChar, adColNullable, 10, , "      ", False, True, , True


DB_CreateField "Transaction_Details", "OrderLineId", adInteger, adColNullable, , , "    ", False, True


DB_CreateField "Transaction_Details", "ScreenPriceOZDolar", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "Transaction_Details", "PerimimDolar", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "Transaction_Details", "GoldGPrice", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "Transaction_Details", "totaNooFgrams", adDouble, adColNullable, , , "      ", False, True
DB_CreateField "Transaction_Details", "PAckingCost", adDouble, adColNullable, , , "      ", False, True

DB_CreateField "Transaction_Details", "ActualGramPrice", adDouble, adColNullable, , , "      ", False, True
 
DB_CreateField "groups", "IsGold", adInteger, adColNullable, , , "    ", False, True

  DB_CreateField "transactionsVatDetails", "Transaction_ID", adDouble, adColNullable, , , , False, True
    
    
    
   DB_CreateField "TblCustemers", "chkTaxExempt", adBoolean, adColNullable, , , "                ", False, True
   
    
        DB_CreateField "TblUsers", "CanChangePriceUpOnly", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "TblUsers", "CanProjectAccountOnly", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "TblUsers", "CanUploadZakat", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "TblUsers", "IsHiddenUser", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "TblUsers", "CanPostPumpInv", adBoolean, adColNullable, , , "", False, True
        
        
        DB_CreateField "tblOPtions", "CanUploadZakatOpt", adBoolean, adColNullable, , , "", False, True
        
   
    DB_CreateField "tblOPtions", "ServerNameW", adVarWChar, adColNullable, 255, , "      ", False
    DB_CreateField "tblOPtions", "DbNameW", adVarWChar, adColNullable, 255, , "      ", False
    
        DB_CreateField "TblOrderUpload", "PostedDate", adDBTimeStamp, adColNullable, , , "      ", False, True
    DB_CreateField "TblOrderUpload", "Approved", adBoolean, adColNullable, , , "???? ?? ??", False, True
      
    DB_CreateField "TblOrderUpload", "Posted", adInteger, adColNullable, , , "      ", False, True

    
    
   DB_CreateField "notes_all", "ItemID", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "notes_all", "UnitID", adInteger, adColNullable, , , "      ", False, True
    
    DB_CreateField "TblClientTransContrDet", "ItemID", adInteger, adColNullable, , , "      ", False, True
    DB_CreateField "TblClientTransContrDet", "UnitID", adInteger, adColNullable, , , "      ", False, True
    
        If DB_CreateTable("tblOilsTypes", True, "ID", False) = True Then
           DB_CreateField "tblOilsTypes", "Name", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblOilsTypes", "NameE", adVarWChar, adColNullable, 255, , "      ", False
            DB_CreateField "tblOilsTypes", "Period", adDouble, adColNullable, , , "  ", False, True
     End If
     DB_CreateField "tblOilsTypes", "KiloMetr", adDouble, adColNullable, , , "  ", False, True
     
     
     
     DB_CreateField "TblRegDateDelgate", "RowId", adGUID, adColNullable, , , "", False, True
     DB_CreateField "TblUsers", "OPenShortInvoicePump", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "OPenShortInvoicePetrol", adBoolean, adColNullable, , , "", False, True
  
   DB_CreateField "TblEmployee", "swapedempid2", adInteger, adColNullable, , , "      ", False, True
   
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 5001 ,'›« Ê—… „” Œ·’ ·„ﬁ«Ê·    ' ,'    Invoice  ' ", "NotesType", 5001
   
   
    DB_CreateField "TblDefComItem", "TransactionID6", adInteger, adColNullable, , , " ???    ", False, True
    
    
    DB_CreateField "TblDefComItem", "NoteSerial16", adVarWChar, adColNullable, 50, , "  ", False, True, , True
    
   
        
        DB_CreateField "TblPayPrePayed", "Account_code", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "TblPayPrePayed", "NCashingType", adInteger, adColNullable, , , """"
        DB_CreateField "Transactions", "TypeTrans", adInteger, adColNullable, , , """"
        
       
       DB_CreateField "DOUBLE_ENTREY_VOUCHERS1", "totalPayed", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblQualityDet", "L", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "W", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "H1", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "H2", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "NoCount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "Width", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "length", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "Height", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblQualityDet", "Area", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TblQualityDet", "OldID", adInteger, adColNullable, , , """"
        DB_CreateField "TblQualityDet", "ColorID", adInteger, adColNullable, , , """"
        DB_CreateField "TblQualityDet", "ItemSize", adInteger, adColNullable, , , """"
        DB_CreateField "TblQualityDet", "ClassId", adInteger, adColNullable, , , """"
        DB_CreateField "TblQualityDet", "LineId", adInteger, adColNullable, , , """"
                        
            

    
    DB_CreateField "projects_des", "QtyOpen", adDouble, adColNullable, , , "    ", False, True
    
    DB_CreateField "TblCustemers", "AddressE", adVarWChar, adColNullable, 400, , "", False, True, , True
    
    
    DB_CreateField "FixedAssets", "BranchName", adVarWChar, adColNullable, 400, , "", False, True, , True
    
DB_CreateField "TblItems", "IsPriceIsLenth", adBoolean, adColNullable, , , "  ", False, True
DB_CreateField "TblItems", "IsPriceIsLenthWH", adBoolean, adColNullable, , , "  ", False, True
      If DB_CreateTable("tblPumpType", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "tblPumpType", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tblPumpType", "BoxId", adInteger, adColNullable, , , ""
        DB_CreateField "tblPumpType", "ItemId", adInteger, adColNullable, , , ""
        DB_CreateField "tblPumpType", "serial", adInteger, adColNullable, , , ""
        DB_CreateField "tblPumpType", "GuaranteeAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblPumpType", "AmountPlus", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblPumpType", "PercentV", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblPumpType", "Total", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "tblPumpType", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblPumpType", "NameE", adVarWChar, adColNullable, 400, , "", False, True, , True
    
    End If
        DB_CreateField "tblPumpType", "Code", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblPumpType", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "tblPumpType", "Name", adVarWChar, adColNullable, 400, , "", False, True, , True

DB_CreateField "Transactions", "OilsTypesID", adInteger, adColNullable, , , ""
  DB_CreateField "tblPumpType", "AmountH", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "tblPumpType", "IsOther", adBoolean, adColNullable, , , "", False, True
  DB_CreateField "tblPumpType", "Account_Code", adVarWChar, adColNullable, 400, , "", False, True, , True
  DB_CreateField "tblPumpType", "Account_CodeComm", adVarWChar, adColNullable, 400, , "", False, True, , True
  
  DB_CreateField "tblPumpType", "StoreID", adInteger, adColNullable, , , ""



DB_CreateField "emp_salary", "TotalVacValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "emp_salary", "vacDay", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "emp_salary", "OverTime", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "emp_salary", "WorkHours", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "emp_salary", "VoCation", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "TblOptions", "DomainData", adLongVarWChar, adColNullable, , , "", False, True, , True

  DB_CreateField "Transaction_Details", "AmountH", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Transaction_Details", "AmountHComm", adDouble, adColNullable, , , "    ", False, True
  DB_CreateField "Transaction_Details", "IsOther", adBoolean, adColNullable, , , "", False, True
  DB_CreateField "Transaction_Details", "Account_Code", adVarWChar, adColNullable, 400, , "", False, True, , True
  DB_CreateField "Transaction_Details", "Account_CodeComm", adVarWChar, adColNullable, 400, , "", False, True, , True
 ' DB_CreateField "Transaction_Details", "DetailsPump", adLongVarWChar, adColNullable, 40000, , "", False, True, , True
  
   DB_CreateField "Transaction_Details", "DetailsPump", adVarWChar, adColNullable, , , "", False, True, , True
  

      If DB_CreateTable("Transaction_DetailsPump", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "Transaction_DetailsPump", "LineID", adInteger, adColNullable, , , ""
        DB_CreateField "Transaction_DetailsPump", "CusID", adInteger, adColNullable, , , ""
        DB_CreateField "Transaction_DetailsPump", "ItemId", adInteger, adColNullable, , , ""

        DB_CreateField "Transaction_DetailsPump", "Amount", adDouble, adColNullable, , , "    ", False, True
        
       
    
    End If
    DB_CreateField "Transaction_DetailsPump", "Qty", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "Transaction_DetailsPump", "Price", adDouble, adColNullable, , , "    ", False, True
    
DB_CreateField "Transaction_DetailsPump", "Transaction_ID", adInteger, adColNullable, , , ""

DB_CreateField "Transaction_DetailsPump", "RecNo", adInteger, adColNullable, , , ""


DB_CreateField "Transactions", "TimeIn", adVarWChar, adColNullable, 50, , "      ", False, True, , True
DB_CreateField "Transactions", "TypeInvoice", adInteger, adColNullable, , , ""
DB_CreateField "Transactions", "ColorID2", adInteger, adColNullable, , , ""

DB_CreateField "Transaction_Details", "PumpId", adInteger, adColNullable, , , ""

DB_CreateField "Transaction_Details", "PrevQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "CurrentQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "UsedQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "StillQty", adDouble, adColNullable, , , "    ", False, True




DB_CreateField "Transaction_Details", "CashQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "MadaQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "VisaQty", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "DeferredQty", adDouble, adColNullable, , , "    ", False, True




DB_CreateField "Transaction_Details", "Cash", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "Mada", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "Visa", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transaction_Details", "Deferred", adDouble, adColNullable, , , "    ", False, True




DB_CreateField "Transactions", "CarCurrentValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CarPrevValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CarEnginoil", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CarPrevValue", adDouble, adColNullable, , , "    ", False, True
DB_CreateField "Transactions", "CarGearOil", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "Transactions", "CarOilChangeDate", adDBTimeStamp, adColNullable, 8, , "", False



    '*************Ship *************
      If DB_CreateTable("TBLProjectBillHistory", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "TBLProjectBillHistory", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLProjectBillHistory", "TblLCID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLProjectBillHistory", "bill_id", adInteger, adColNullable, , , ""
        DB_CreateField "TBLProjectBillHistory", "serial", adInteger, adColNullable, , , ""
        DB_CreateField "TBLProjectBillHistory", "GuaranteeAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLProjectBillHistory", "AmountPlus", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLProjectBillHistory", "AmountMin", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLProjectBillHistory", "Total", adDouble, adColNullable, , , "    ", False, True
        
        
    
    End If
    
    
    DB_CreateField "project_billl", "TotalBefore", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "project_billl", "Discount4", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "project_billl", "Discount3", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "project_billl", "BondAmt", adDouble, adColNullable, , , "    ", False, True


    

    
    If DB_CreateTable("tblCarMaint", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "tblCarMaint", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "tblCarMaint", "RecordDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "tblCarMaint", "NoteID", adInteger, adColNullable, , , ""
        DB_CreateField "tblCarMaint", "GroupId", adInteger, adColNullable, , , ""
        DB_CreateField "tblCarMaint", "ModelId", adInteger, adColNullable, , , ""
        DB_CreateField "tblCarMaint", "ItemId", adInteger, adColNullable, , , ""
        DB_CreateField "tblCarMaint", "ChkOrg", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "tblCarMaint", "chkCom", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "tblCarMaint", "chkTested", adBoolean, adColNullable, , , "", False, True
        DB_CreateField "tblCarMaint", "chkNormal", adBoolean, adColNullable, , , "", False, True

        
        
    End If
                    
              
                 
DB_CreateField "Transactions", "YearFact", adInteger, adColNullable, , , "    ", False, True
      DB_CreateField "Transactions", "CarTypeID", adInteger, adColNullable, , , "    ", False, True
      
      DB_CreateField "Transactions", "Shaseh", adVarWChar, adColNullable, 255, , "", False, True, , True
      DB_CreateField "Transactions", "CarMeter", adVarWChar, adColNullable, 255, , "", False, True, , True
      DB_CreateField "Transactions", "PlateNo", adVarWChar, adColNullable, 255, , "", False, True, , True
                
      
      DB_CreateField "Transactions", "MrNo", adVarWChar, adColNullable, 255, , "", False, True, , True
      
     DB_CreateField "TblEmpAllocations", "FromDate", adDBTimeStamp, adColNullable, 8, , "", False
     DB_CreateField "TblEmpAllocations", "ToDate", adDBTimeStamp, adColNullable, 8, , "", False
    
   DB_CreateField "TblBillComputerChekDetails", "MainType", adVarWChar, adColNullable, 400, , "", False, True, , True
    
    DB_CreateField "TblBillComputerChek", "subcar1", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar2", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar3", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar4", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar5", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar6", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar7", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar8", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar9", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar10", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar11", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar12", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar13", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblBillComputerChek", "subcar14", adBoolean, adColNullable, , , "", False, True
    
    
    DB_CreateField "TblComputerChek", "IsMain", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblComputerChek", "MainID", adInteger, adColNullable, , , ""
    DB_CreateField "transactions", "ExcelRow", adInteger, adColNullable, , , ""
    DB_CreateField "transactions", "ExcelFile", adVarWChar, adColNullable, 400, , "", False, True, , True
    
    DB_CreateField "TBLLCHistory", "MarginNo", adInteger, adColNullable, , , ""
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22004 ,'Margin advice' ,'       Project' ", "NotesType", 22004
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22006 ,'Margin advice' ,'       Project' ", "NotesType", 22006
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22007 ,'Margin advice' ,'       Project' ", "NotesType", 22007
   
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22008 ,'Margin advice' ,'       Project' ", "NotesType", 22008
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22009 ,'Margin advice' ,'       Project' ", "NotesType", 22009
   
   add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 22010 ,'Open' ,'       Project' ", "NotesType", 22010
   
    
        DB_CreateField "TBLLCHistory", "NoteID", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLCHistory", "NoteSerial", adInteger, adColNullable, , , ""
    
         DB_CreateField "TblFiterWaiver", "DaysValueIncrease", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DaysValueIncomplete", adDouble, adColNullable, , , " ", False, True


    DB_CreateField "tblVacancy", "Account_Code", adVarWChar, adColNullable, 255, , "", False, True, , True
    
    
        DB_CreateField "TblFiterWaiverDet2", "RemainCommissions", adDouble, adColNullable, , , "      ", False, True
      
  DB_CreateField "TblOtheExpensAqar", "IsLegalAffairs", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblOtheExpensAqar", "LegalAffairs", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TblOtheExpensAqar", "LegalAffairsDate", adDBTimeStamp, adColNullable, 8, , "", False
    
DB_CreateField "TblEmpDepartments", "MokafahVacID", adInteger, adColNullable, , , ""
    DB_CreateField "TblFiterWaiver", "DayValueInc", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayCountInc", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayValueIncomplete", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiver", "DayCountIncomplete", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiverDet2", "TotalStill", adDouble, adColNullable, , , " ", False, True
    DB_CreateField "TblFiterWaiverDet2", "RemainCommissions", adDouble, adColNullable, , , " ", False, True

  
  DB_CreateField "TblFiterWaiver", "IsLegalAffairs", adBoolean, adColNullable, , , "", False, True
DB_CreateField "TblFiterWaiver", "LegalAffairs", adVarWChar, adColNullable, 400, , "", False, True, , True

     
     DB_CreateField "LCTypes", "TypeLCLG", adInteger, adColNullable, , , ""
    DB_CreateField "TblUsers", "CantWorkwithComponenetinEmpScr", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "AllowEditCreditLimit", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "AllowEditCreditBalance", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "AllowConvertAlertToJob", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "AllowSkipDiscountGroup", adBoolean, adColNullable, , , "", False, True
    DB_CreateField "TblUsers", "HideInfroCasher", adBoolean, adColNullable, , , "", False, True
    
    DB_CreateField "TblUsers", "CanEditLegalAffairs", adBoolean, adColNullable, , , "", False, True
    
    DB_CreateField "tblVacancy", "ComponentID", adInteger, adColNullable, , , ""
   '*************Ship *************
      If DB_CreateTable("TBLLCMargin", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "TBLLCMargin", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin", "TblLCID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin", "serial", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin", "MarginNo", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin", "GuaranteeDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin", "Amount", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TBLLCMargin", "MarginAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
        DB_CreateField "TBLLCMargin", "BankAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
        
        
        
    
    End If
    DB_CreateField "TBLLCMargin", "BankAccountCode2", adVarWChar, adColNullable, 255, , "", False, True, , True
    DB_CreateField "TBLLCMargin", "AccountMargen2", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TBLLCMargin", "MargenValue", adDouble, adColNullable, , , "    ", False, True




DB_CreateField "TBLLCMargin", "AccountMargen2", adVarWChar, adColNullable, 400, , "", False, True, , True
DB_CreateField "TBLLCMargin", "MargenValue", adDouble, adColNullable, , , "    ", False, True

DB_CreateField "TBLLCMargin", "OrderDate", adDBTimeStamp, adColNullable, 8, , "", False
    DB_CreateField "TBLLCMargin", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin", "NoteID2", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLCMargin", "NoteSerial2", adInteger, adColNullable, , , ""
    
    DB_CreateField "TBLLCMargin", "Type", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLCMargin", "PayedAmount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLCMargin", "StillAmount", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TBLLCMargin", "NoteID", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLCMargin", "NoteSerial", adInteger, adColNullable, , , ""
    
       DB_CreateField "TBLLCMargin", "NoteID2", adInteger, adColNullable, , , ""
    DB_CreateField "TBLLCMargin", "NoteSerial2", adInteger, adColNullable, , , ""
 DB_CreateField "TBLLCMargin", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin", "IsFullPayed", adBoolean, adColNullable, , , "", False, True
        
    
    
    
    
    
     If DB_CreateTable("TBLLCMargin2", True, "ID", True) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "TBLLCMargin2", "ID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "TblLCID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "serial", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "MarginNo", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "GuaranteeDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin2", "Amount", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TBLLCMargin2", "MarginAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
        DB_CreateField "TBLLCMargin2", "BankAccountCode", adVarWChar, adColNullable, 255, , "", False, True, , True
        
        DB_CreateField "TBLLCMargin2", "AccountMargen2", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "TBLLCMargin2", "MargenValue", adDouble, adColNullable, , , "    ", False, True
        
        
        
        
        DB_CreateField "TBLLCMargin2", "AccountMargen2", adVarWChar, adColNullable, 400, , "", False, True, , True
        DB_CreateField "TBLLCMargin2", "MargenValue", adDouble, adColNullable, , , "    ", False, True
        
        DB_CreateField "TBLLCMargin2", "OrderDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin2", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin2", "NoteID2", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "NoteSerial2", adInteger, adColNullable, , , ""
        
        DB_CreateField "TBLLCMargin2", "Type", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "PayedAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLLCMargin2", "StillAmount", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TBLLCMargin2", "NoteID", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "NoteSerial", adInteger, adColNullable, , , ""
        
        DB_CreateField "TBLLCMargin2", "NoteID2", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "NoteSerial2", adInteger, adColNullable, , , ""
        DB_CreateField "TBLLCMargin2", "PayDate", adDBTimeStamp, adColNullable, 8, , "", False
        DB_CreateField "TBLLCMargin2", "IsFullPayed", adBoolean, adColNullable, , , "", False, True
        
    
    End If
    
    
    DB_CreateField "TBLLCMargin2", "IsOpenBalance", adBoolean, adColNullable, , , "", False, True
            
            
            
            DB_CreateField "TBLLCMargin2", "RowId", adGUID, adColNullable, , , "", False, True
            DB_CreateField "TBLLCMargin", "RowId", adGUID, adColNullable, , , "", False, True
            DB_CreateField "tblLCOpenB", "RowId", adGUID, adColNullable, , , "", False, True
            
            DB_CreateField "tblLC", "NoteIDOpenRowId", adGUID, adColNullable, , , "", False, True
            DB_CreateField "tblLC", "NoteID2RowId", adGUID, adColNullable, , , "", False, True
            DB_CreateField "tblLC", "NoteIDRowId", adGUID, adColNullable, , , "", False, True
            
            
            DB_CreateField "Notes", "RowId", adGUID, adColNullable, , , "", False, True
            DB_CreateField "Notes1", "RowId", adGUID, adColNullable, , , "", False, True
    


    
    
    Update30Follow
    '*****************
    '*****************************

        
    



End Function




 
Function updatedbExamples()
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

    '„À«· ⁄·Ï «‰‘«¡ ÃœÊ·
    If DB_CreateTable("Ahmed", True, "ID_man", False) = True Then
    
        '„À«· ⁄·Ï «‰‘«¡ Õﬁ·
        DB_CreateField "Ahmed", "name", adVarWChar, adColNullable, 255, , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True, , True
        DB_CreateField "Ahmed", "sex", adBinary, adColNullable, , , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True

        DB_CreateField "Ahmed", "date4", adDBTimeStamp, adColNullable, 8, , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False
    End If

    If DB_CreateTable("salim", True, "ID_man1") = True Then
   
        DB_CreateField "salim", "name", adVarWChar, adColNullable, 255, , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True, , True
        DB_CreateField "salim", "sex", adBinary, adColNullable, , , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True

        DB_CreateField "salim", "id", adInteger, adColNullable, , , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True
    End If

    'db_createRelationSQL "ahmed", "ID_man", "salim", "id"
 
    Exit Function
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '„À«· ⁄·Ï «÷«›Â Õﬁ· ·ÃœÊ·
    DB_CreateField "currency", "divname", adVarWChar, adColNullable, 255, , "ﬂ”— «·⁄„·Â", False, True
    DB_CreateField "currency", "divnamee", adVarWChar, adColNullable, 255, , "ﬂ”— «·⁄„·Â «‰Ã·Ì“Ì", False, True
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '„À«· €·Ï  ⁄œÌ· ‰Ê⁄ Õﬁ·
    DB_updateField "ahmed", "sex ", "nvarchar(255) not null  "
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '„À«· ⁄·Ï «÷«›… ”Ã· ·ÃœÊ·
    add_record_to_table "ahmed", "name,sex,date4", "'ali samy',1,'11-2-1980'"
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

    '„À«· ⁄·Ï «‰‘«¡ ⁄·«ﬁ…
    db_createRelationSQL "ahmed", "ID_man", "salim", "id"
    '„À«· ⁄·Ï Õ–› ⁄·«ﬁ…
    db_deleteRelationSQL "ahmed", "ID_man", "salim", "id"
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

    Dim New_View As String
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '„À«· ⁄·Ï «‰‘«¡ «” ⁄·«„ «Ê  ⁄œÌ· «” ⁄·«„ „ÊÃÊœ
    New_View = " SELECT      dbo.Notes.Note_Value,dbo.Notes.NoteSerial1,dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, " & "    dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], " & "   dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description " & "FROM         dbo.Notes INNER JOIN " & "                     dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID "

    db_createOrUpdateviewSQL "ahmed_view", New_View

End Function

 
 


Sub CreateOrUpdateTrigger()
    Dim Cmd As ADODB.Command
    Dim triggerName As String
    Dim checkTriggerSQL As String
    Dim createTriggerSQL As String

    ' «”„ «· —ÌÃ—
    triggerName = "trg_update_double_entry_vouchers"

    ' SQL ·· Õﬁﬁ ≈–« ﬂ«‰ «· —ÌÃ— „ÊÃÊœ«° Ê≈–« ﬂ«‰ „ÊÃÊœ« Ì „ Õ–›Â
    checkTriggerSQL = "IF EXISTS (SELECT * FROM sys.triggers WHERE name = '" & triggerName & "') " & _
                      "BEGIN " & _
                      "DROP TRIGGER " & triggerName & " " & _
                      "END;"

    ' SQL ·≈‰‘«¡ «· —ÌÃ—
    createTriggerSQL = "CREATE TRIGGER " & triggerName & " " & _
                       "ON project_billl " & _
                       "AFTER UPDATE " & _
                       "AS " & _
                       "BEGIN " & _
                       "IF UPDATE(approved) " & _
                       "BEGIN " & _
                       "UPDATE DOUBLE_ENTREY_VOUCHERS " & _
                       "SET Posted = NULL " & _
                       "FROM DOUBLE_ENTREY_VOUCHERS DE " & _
                       "INNER JOIN inserted i ON DE.Notes_ID = i.note_id " & _
                       "WHERE i.Approved = 1; " & _
                       "END " & _
                       "END;"

    ' ≈‰‘«¡ ﬂ«∆‰ «·√„—
    Set Cmd = New ADODB.Command
    Set Cmd.ActiveConnection = Cn ' «” Œœ«„ «·« ’«· «·„› ÊÕ cn

    '  ‰›Ì– SQL ·· Õﬁﬁ „‰ ÊÃÊœ «· —ÌÃ— ÊÕ–›Â ≈–« ﬂ«‰ „ÊÃÊœ«
    Cmd.CommandText = checkTriggerSQL
    Cmd.Execute

    '  ‰›Ì– SQL ·≈‰‘«¡ «· —ÌÃ—
    Cmd.CommandText = createTriggerSQL
    Cmd.Execute

    '  ‰ŸÌ›
    Set Cmd = Nothing

  '  MsgBox " „ ≈‰‘«¡ √Ê  ⁄œÌ· «· —ÌÃ— »‰Ã«Õ."
End Sub


'==== helpers =========================================
Private Function GetColumnType(ByVal tablename As String, ByVal ColName As String) As String
    Dim rs As New ADODB.Recordset, sql As String
    sql = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS " & _
          "WHERE TABLE_NAME='" & tablename & "' AND COLUMN_NAME='" & ColName & "'"
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then GetColumnType = LCase$(rs!DATA_TYPE & "")
    rs.Close: Set rs = Nothing
End Function

Private Function ColumnExists(ByVal tablename As String, ByVal ColName As String) As Boolean
    Dim rs As New ADODB.Recordset, sql As String
    sql = "SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS " & _
          "WHERE TABLE_NAME='" & tablename & "' AND COLUMN_NAME='" & ColName & "'"
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
    ColumnExists = Not rs.EOF
    rs.Close: Set rs = Nothing
End Function

Private Function TableRowCount(ByVal tablename As String) As Long
    Dim rs As New ADODB.Recordset, sql As String
    sql = "SELECT COUNT(*) AS Cnt FROM " & tablename
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then TableRowCount = CLng(val(rs!cnt & ""))
    rs.Close: Set rs = Nothing
End Function

'==== core per-table convert ===========================
Private Sub ConvertInvoiceIdForTable(ByVal tablename As String)
    On Error GoTo eh

    Dim colType As String
    Dim cnt As Long
    Dim sql As String

    colType = GetColumnType(tablename, "InvoiceID")
    cnt = TableRowCount(tablename)

    If colType <> "float" Then
        ' „‘ Float ? „›Ì‘ Õ«Ã… ‰⁄„·Â«
        Exit Sub
    End If
    If cnt <> 0 Then
        ' «·ÃœÊ· ›ÌÂ »Ì«‰«  ? Õ›«Ÿ« ⁄·Ï «·œ« « ·‰ ‰‰›–
        Exit Sub
    End If

    Cn.BeginTrans

    ' 1) √÷› ⁄„Êœ „ƒﬁ  ‰’Ì ≈‰ ·„ Ìﬂ‰ „ÊÃÊœ«
    If Not ColumnExists(tablename, "InvoiceID_str") Then
        sql = "ALTER TABLE " & tablename & " ADD InvoiceID_str NVARCHAR(50) NULL"
        Cn.Execute sql
    End If

    ' 2) «„·√Â » ÕÊÌ· ’ÕÌÕ: FLOAT -> DECIMAL(38,0) -> NVARCHAR
    sql = "UPDATE " & tablename & " " & _
          "SET InvoiceID_str = CONVERT(NVARCHAR(50), CAST(InvoiceID AS DECIMAL(38,0)))"
    Cn.Execute sql

    ' 3) «Õ–› «·⁄„Êœ «·ﬁœÌ„
    sql = "ALTER TABLE " & tablename & " DROP COLUMN InvoiceID"
    Cn.Execute sql

    ' 4) √⁄œ  ”„Ì… «·⁄„Êœ «·ÃœÌœ ≈·Ï «·«”„ «·√’·Ì
    sql = "EXEC sp_rename '" & tablename & ".InvoiceID_str','InvoiceID','COLUMN'"
    Cn.Execute sql

    Cn.CommitTrans
    Exit Sub

eh:
    On Error Resume Next
    Cn.RollbackTrans
    MsgBox "›‘·  ÕÊÌ· " & tablename & ".InvoiceID" & vbCrLf & Err.Description, vbExclamation
End Sub

'==== run for both tables ==============================
Public Sub ConvertInvoiceIdAll()
    ConvertInvoiceIdForTable "tblEInvoice"
    ConvertInvoiceIdForTable "tblEInvoice2"
    ConvertInvoiceIdForTable "tmptblEInvoice"
    
    'MsgBox " „  «·„⁄«·Ã… (≈‰  Ê›—  «·‘—Êÿ) ··ÃœÊ·Ì‰.", vbInformation
End Sub


