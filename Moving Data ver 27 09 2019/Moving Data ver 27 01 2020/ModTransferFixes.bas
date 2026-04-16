Attribute VB_Name = "ModTransferFixes"
Option Explicit

Private gTransferLastSQL As String
Private Const TRANSFER_BATCH_SIZE As Long = 50
Private Const TRANSFER_VALIDATE_EVERY As Long = 250

Private Type TTransferCounters
    Inserted As Long
    Updated As Long
    Failed As Long
    Skipped As Long
End Type

Private Type TTransferRun
    SessionCode As String
    StartedAt As Date
    FirstError As String
    FirstErrorSQL As String
    LogFile As String
End Type

Public Function TransferItemsFromServerToBranch(ByRef UserMessage As String, Optional ByRef TraceText As String = "") As Boolean
    On Error GoTo EH

    Dim runInfo As TTransferRun
    Dim totalIns As Long
    Dim totalUpd As Long
    Dim totalFail As Long
    Dim totalSkip As Long
    Dim cnt As TTransferCounters
    Dim inTx As Boolean
    Dim s As String

    runInfo = TransferRunStart("Items_ServerToPOS")

    TransferWriteLog runInfo, "START TransferItemsFromServerToBranch"
    TransferValidateReady runInfo

    POSConnection.BeginTrans
    inTx = True
    TransferWriteLog runInfo, "POS transaction started"

    cnt = SyncMissingRows(runInfo, _
                          "TblItems", _
                          "SELECT ItemID, ItemCode, ItemName, DefaultSupplier, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, itemSerials, barCodeNO, SizeID11 FROM TblItems ORDER BY ItemID", _
                          "ItemID", _
                          "ItemID,ItemCode,ItemName,DefaultSupplier,GroupID,HaveSerial,LastUpdate,PurchasePrice,SallingPrice,RequestLimit,CustomerPrice,HaveGuarantee,GuaranteeValue,GuaranteeType,IsArchive,ItemType,AssbliedItem,RelatedItem,ItemComment,ItemCase,ItemMaking,ItemMakingNew,code,Branch_NO,Fullcode,prifix,PartNo,CostPrice,ItemNamee,itemSerials,barCodeNO,SizeID11")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblItemsUnits", _
                          "SELECT JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2 FROM TblItemsUnits ORDER BY ItemID, UnitID", _
                          "ItemID,UnitID", _
                          "JunckID,ItemID,UnitID,UnitFactor,SecOrder,DefaultUnit,UnitSalesPrice,UnitPurPrice,FactorByDefaultUnit,MinSelingPrice,ForUnit,MethodCalc,SessionCode,barCodeNo2")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncSingleRowUpdate(runInfo, _
                              "TblOptions", _
                              "SELECT TOP 1 BigUserPw, BigUserPw2 FROM TblOptions", _
                              "BigUserPw,BigUserPw2", _
                              "")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    POSConnection.CommitTrans
    inTx = False
    TransferWriteLog runInfo, "POS transaction committed successfully"

    s = "Items were transferred successfully from server to branch." & vbCrLf & _
        "Inserted: " & CStr(totalIns) & vbCrLf & _
        "Updated: " & CStr(totalUpd) & vbCrLf & _
        "Skipped: " & CStr(totalSkip) & vbCrLf & _
        "Failed: " & CStr(totalFail) & vbCrLf & _
        "Trace file: " & runInfo.LogFile

    UserMessage = s
    TraceText = s
    TransferItemsFromServerToBranch = True
    Exit Function

EH:
    If inTx Then
        On Error Resume Next
        POSConnection.RollbackTrans
        On Error GoTo 0
        TransferWriteLog runInfo, "POS transaction rolled back"
    End If

    TransferRememberError runInfo, Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR TransferItemsFromServerToBranch: " & Err.Number & " - " & Err.Description

    UserMessage = BuildTransferFailureMessage(runInfo, "Failed to transfer items from server to branch")
    TraceText = UserMessage
    TransferItemsFromServerToBranch = False
End Function

Public Function UpdatePricesFromServerToBranch(ByRef UserMessage As String, Optional ByRef TraceText As String = "") As Boolean
    On Error GoTo EH

    Dim runInfo As TTransferRun
    Dim totalIns As Long
    Dim totalUpd As Long
    Dim totalFail As Long
    Dim totalSkip As Long
    Dim cnt As TTransferCounters
    Dim inTx As Boolean
    Dim s As String

    runInfo = TransferRunStart("CoreAndPrices_ServerToPOS")

    TransferWriteLog runInfo, "START UpdatePricesFromServerToBranch"
    TransferValidateReady runInfo

    POSConnection.BeginTrans
    inTx = True
    TransferWriteLog runInfo, "POS transaction started"

    cnt = SyncMissingRows(runInfo, "Groups", "SELECT GroupID, GroupName FROM Groups ORDER BY GroupID", "GroupID", "GroupID,GroupName")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, "TblUnites", "SELECT UnitID, UnitName, UnitNamee FROM TblUnites ORDER BY UnitID", "UnitID", "UnitID,UnitName,UnitNamee")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblItems", _
                          "SELECT ItemID, ItemCode, ItemName, GroupID, HaveSerial, LastUpdate, PurchasePrice, SallingPrice, RequestLimit, CustomerPrice, HaveGuarantee, GuaranteeValue, GuaranteeType, IsArchive, ItemType, AssbliedItem, RelatedItem, ItemComment, ItemCase, ItemMaking, ItemMakingNew, code, Branch_NO, Fullcode, prifix, PartNo, CostPrice, ItemNamee, DefaultSupplier, itemSerials, barCodeNO, SizeID11 FROM TblItems ORDER BY ItemID", _
                          "ItemID", _
                          "ItemID,ItemCode,ItemName,GroupID,HaveSerial,LastUpdate,PurchasePrice,SallingPrice,RequestLimit,CustomerPrice,HaveGuarantee,GuaranteeValue,GuaranteeType,IsArchive,ItemType,AssbliedItem,RelatedItem,ItemComment,ItemCase,ItemMaking,ItemMakingNew,code,Branch_NO,Fullcode,prifix,PartNo,CostPrice,ItemNamee,DefaultSupplier,itemSerials,barCodeNO,SizeID11")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblPaymentType", _
                          "SELECT PaymentID, PaymentName, PaymentNamee, Accountcom, commision, branch_no, TaxTobacco, AccTaxTobacco, IsNewCode, IsHiddenVat, IsDefault FROM TblPaymentType ORDER BY PaymentID", _
                          "PaymentID", _
                          "PaymentID,PaymentName,PaymentNamee,Accountcom,commision,branch_no,TaxTobacco,AccTaxTobacco,IsNewCode,IsHiddenVat,IsDefault")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, "TblPaymentUser", "SELECT PaynetID, UserID FROM TblPaymentUser ORDER BY PaynetID, UserID", "PaynetID,UserID", "PaynetID,UserID")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "BanksData", _
                          "SELECT BankID, BankName, BankNamee, Account_Code, Account_Code1, Account_Code2, BranchId, ParetnAccount, parent_account FROM BanksData ORDER BY BankID", _
                          "BankID", _
                          "BankID,BankName,BankNamee,Account_Code,Account_Code1,Account_Code2,BranchId,ParetnAccount,parent_account")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblUsers", _
                          "SELECT UserID, UserName, UserPassword, UserLevel, UserSign, AllowDelete, EmployeeID, UserNamee, IsConnectedOnline FROM TblUsers ORDER BY UserID", _
                          "UserID", _
                          "UserID,UserName,UserPassword,UserLevel,UserSign,AllowDelete,EmployeeID,UserNamee,IsConnectedOnline")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblEmpJobsTypes", _
                          "SELECT JobID, JobName, JobNamee FROM TblEmpJobsTypes ORDER BY JobID", _
                          "JobID", _
                          "JobID,JobName,JobNamee")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblEmployee", _
                          "SELECT EmpID, EmpName, JobID, SalValue, Mobile, NationalNo, Address, EmpNamee, BranchID, UserName, UserPw FROM TblEmployee ORDER BY EmpID", _
                          "EmpID", _
                          "EmpID,EmpName,JobID,SalValue,Mobile,NationalNo,Address,EmpNamee,BranchID,UserName,UserPw")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblCustemers", _
                          "SELECT CusID, CusName, CusNamee, ResponsibleContact, Cus_mobile, Type, OpenBalance, Account_Code, CityID, EmpId, Address, parent_account, prifix, Fullcode, BranchId, VATNO, CustGID FROM TblCustemers ORDER BY CusID", _
                          "CusID", _
                          "CusID,CusName,CusNamee,ResponsibleContact,Cus_mobile,Type,OpenBalance,Account_Code,CityID,EmpId,Address,parent_account,prifix,Fullcode,BranchId,VATNO,CustGID")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblStore", _
                          "SELECT StoreID, StoreName, Account_Code, Account_Code1, Account_Code2, Emp_ID, Account_Code3, linked, BranchId, Code, StoreNamee, ParetnAccount, SalesPersonId, PurchasePersonid, Account_Code0, Account_Code11, Account_Code22, Account_Code33, BoxID FROM TblStore ORDER BY StoreID", _
                          "StoreID", _
                          "StoreID,StoreName,Account_Code,Account_Code1,Account_Code2,Emp_ID,Account_Code3,linked,BranchId,Code,StoreNamee,ParetnAccount,SalesPersonId,PurchasePersonid,Account_Code0,Account_Code11,Account_Code22,Account_Code33,BoxID")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMissingRows(runInfo, _
                          "TblItemsUnits", _
                          "SELECT JunckID, ItemID, UnitID, UnitFactor, SecOrder, DefaultUnit, UnitSalesPrice, UnitPurPrice, FactorByDefaultUnit, MinSelingPrice, ForUnit, MethodCalc, SessionCode, barCodeNo2 FROM TblItemsUnits ORDER BY ItemID, UnitID", _
                          "ItemID,UnitID", _
                          "JunckID,ItemID,UnitID,UnitFactor,SecOrder,DefaultUnit,UnitSalesPrice,UnitPurPrice,FactorByDefaultUnit,MinSelingPrice,ForUnit,MethodCalc,SessionCode,barCodeNo2")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMatchedRows(runInfo, _
                          "TblItemsUnits", _
                          "SELECT ItemID, UnitID, UnitSalesPrice, barCodeNo2, MaxSelingPrice, UnitWholeSalePrice, MinSelingPrice, UnitPurPrice FROM TblItemsUnits ORDER BY ItemID, UnitID", _
                          "ItemID,UnitID", _
                          "UnitSalesPrice,barCodeNo2,MaxSelingPrice,UnitWholeSalePrice,MinSelingPrice,UnitPurPrice")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    cnt = SyncMatchedRows(runInfo, _
                          "TblItems", _
                          "SELECT ItemID, ItemName, barCodeNO, Code, Fullcode, IsArchive FROM TblItems ORDER BY ItemID", _
                          "ItemID", _
                          "ItemName,barCodeNO,Code,Fullcode,IsArchive")
    AccumulateCounters cnt, totalIns, totalUpd, totalFail, totalSkip

    POSConnection.CommitTrans
    inTx = False
    TransferWriteLog runInfo, "POS transaction committed successfully"

    s = "Prices and master data were updated successfully from server to branch." & vbCrLf & _
        "Inserted: " & CStr(totalIns) & vbCrLf & _
        "Updated: " & CStr(totalUpd) & vbCrLf & _
        "Skipped: " & CStr(totalSkip) & vbCrLf & _
        "Failed: " & CStr(totalFail) & vbCrLf & _
        "Trace file: " & runInfo.LogFile

    UserMessage = s
    TraceText = s
    UpdatePricesFromServerToBranch = True
    Exit Function

EH:
    If inTx Then
        On Error Resume Next
        POSConnection.RollbackTrans
        On Error GoTo 0
        TransferWriteLog runInfo, "POS transaction rolled back"
    End If

    TransferRememberError runInfo, Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR UpdatePricesFromServerToBranch: " & Err.Number & " - " & Err.Description

    UserMessage = BuildTransferFailureMessage(runInfo, "Failed to update prices and master data")
    TraceText = UserMessage
    UpdatePricesFromServerToBranch = False
End Function

Private Sub TransferValidateReady(ByRef runInfo As TTransferRun)
    On Error GoTo EH

    TransferEnsureConnectionState Cn, "Server connection", runInfo, "Transfer startup"
    TransferEnsureConnectionState POSConnection, "Branch connection", runInfo, "Transfer startup"

    TransferWriteLog runInfo, "Server connection OK"
    TransferWriteLog runInfo, "POS connection OK"
    Exit Sub

EH:
    TransferRememberError runInfo, Err.Description, ""
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function SyncMissingRows(ByRef runInfo As TTransferRun, ByVal TableName As String, ByVal SourceSQL As String, ByVal KeyFieldsCsv As String, ByVal InsertFieldsCsv As String) As TTransferCounters
    On Error GoTo EH

    Dim rsSrc As ADODB.Recordset
    Dim rsSchema As ADODB.Recordset
    Dim existingKeys As Collection
    Dim keys() As String
    Dim fields() As String
    Dim rowCount As Long
    Dim sqlInsert As String
    Dim rowKey As String
    Dim batchSQL As String
    Dim batchCount As Long

    keys = CsvToArray(KeyFieldsCsv)
    fields = CsvToArray(InsertFieldsCsv)

    TransferWriteLog runInfo, "SyncMissingRows begin: " & TableName
    TransferEnsureConnectionState Cn, "Server connection", runInfo, TableName & " source open"
    TransferEnsureConnectionState POSConnection, "Branch connection", runInfo, TableName & " key load"
    TransferValidateFields Cn, TableName, fields, runInfo, "Source"
    TransferValidateFields POSConnection, TableName, fields, runInfo, "Destination"
    Set existingKeys = LoadExistingKeys(POSConnection, TableName, keys, runInfo)

    Set rsSchema = OpenSchemaRecordset(Cn, TableName)
    Set rsSrc = New ADODB.Recordset
    rsSrc.Open SourceSQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsSrc.EOF
        rowCount = rowCount + 1
        If (rowCount Mod TRANSFER_VALIDATE_EVERY) = 0 Then
            TransferValidateLoopConnections runInfo, TableName & " insert scan", rowCount
        End If

        rowKey = BuildCompositeKeyFromRecordset(rsSrc, keys)
        If CollectionContainsKey(existingKeys, rowKey) Then
            SyncMissingRows.Skipped = SyncMissingRows.Skipped + 1
        Else
            sqlInsert = BuildInsertSQLFromRS(TableName, rsSrc, fields, rsSchema)
            QueueBatchStatement batchSQL, batchCount, sqlInsert
            AddCollectionKey existingKeys, rowKey
            If batchCount >= TRANSFER_BATCH_SIZE Then
                FlushQueuedBatch POSConnection, batchSQL, batchCount, SyncMissingRows.Inserted, runInfo, TableName & " INSERT"
            End If
        End If
        rsSrc.MoveNext
    Loop

    FlushQueuedBatch POSConnection, batchSQL, batchCount, SyncMissingRows.Inserted, runInfo, TableName & " INSERT"

    TransferWriteLog runInfo, TableName & " source rows read: " & CStr(rowCount)
    TransferWriteLog runInfo, TableName & " inserted: " & CStr(SyncMissingRows.Inserted) & ", skipped: " & CStr(SyncMissingRows.Skipped)

CleanExit:
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Exit Function

EH:
    SyncMissingRows.Failed = SyncMissingRows.Failed + 1
    TransferRememberError runInfo, "Table=" & TableName & " - " & Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR SyncMissingRows " & TableName & ": " & Err.Description
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Err.Raise Err.Number, IIf(Len(Err.Source) = 0, "ModTransferFixes.SyncMissingRows", Err.Source), Err.Description
End Function

Private Function SyncMatchedRows(ByRef runInfo As TTransferRun, ByVal TableName As String, ByVal SourceSQL As String, ByVal KeyFieldsCsv As String, ByVal UpdateFieldsCsv As String) As TTransferCounters
    On Error GoTo EH

    Dim rsSrc As ADODB.Recordset
    Dim rsSchema As ADODB.Recordset
    Dim existingKeys As Collection
    Dim keys() As String
    Dim fields() As String
    Dim rowCount As Long
    Dim sqlUpdate As String
    Dim rowKey As String
    Dim batchSQL As String
    Dim batchCount As Long

    keys = CsvToArray(KeyFieldsCsv)
    fields = CsvToArray(UpdateFieldsCsv)

    TransferWriteLog runInfo, "SyncMatchedRows begin: " & TableName
    TransferEnsureConnectionState Cn, "Server connection", runInfo, TableName & " source open"
    TransferEnsureConnectionState POSConnection, "Branch connection", runInfo, TableName & " key load"
    TransferValidateFields Cn, TableName, JoinArray(keys, fields), runInfo, "Source"
    TransferValidateFields POSConnection, TableName, JoinArray(keys, fields), runInfo, "Destination"
    Set existingKeys = LoadExistingKeys(POSConnection, TableName, keys, runInfo)

    Set rsSchema = OpenSchemaRecordset(Cn, TableName)
    Set rsSrc = New ADODB.Recordset
    rsSrc.Open SourceSQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsSrc.EOF
        rowCount = rowCount + 1
        If (rowCount Mod TRANSFER_VALIDATE_EVERY) = 0 Then
            TransferValidateLoopConnections runInfo, TableName & " update scan", rowCount
        End If

        rowKey = BuildCompositeKeyFromRecordset(rsSrc, keys)
        If CollectionContainsKey(existingKeys, rowKey) Then
            sqlUpdate = BuildUpdateSQLFromRS(TableName, rsSrc, fields, keys, rsSchema)
            QueueBatchStatement batchSQL, batchCount, sqlUpdate
            If batchCount >= TRANSFER_BATCH_SIZE Then
                FlushQueuedBatch POSConnection, batchSQL, batchCount, SyncMatchedRows.Updated, runInfo, TableName & " UPDATE"
            End If
        Else
            SyncMatchedRows.Skipped = SyncMatchedRows.Skipped + 1
        End If
        rsSrc.MoveNext
    Loop

    FlushQueuedBatch POSConnection, batchSQL, batchCount, SyncMatchedRows.Updated, runInfo, TableName & " UPDATE"

    TransferWriteLog runInfo, TableName & " matched rows read: " & CStr(rowCount)
    TransferWriteLog runInfo, TableName & " updated: " & CStr(SyncMatchedRows.Updated) & ", skipped: " & CStr(SyncMatchedRows.Skipped)

CleanExit:
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Exit Function

EH:
    SyncMatchedRows.Failed = SyncMatchedRows.Failed + 1
    TransferRememberError runInfo, "Table=" & TableName & " - " & Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR SyncMatchedRows " & TableName & ": " & Err.Description
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Err.Raise Err.Number, IIf(Len(Err.Source) = 0, "ModTransferFixes.SyncMatchedRows", Err.Source), Err.Description
End Function

Private Function SyncSingleRowUpdate(ByRef runInfo As TTransferRun, ByVal TableName As String, ByVal SourceSQL As String, ByVal UpdateFieldsCsv As String, ByVal WhereClause As String) As TTransferCounters
    On Error GoTo EH

    Dim rsSrc As ADODB.Recordset
    Dim rsSchema As ADODB.Recordset
    Dim fields() As String
    Dim sqlUpdate As String

    fields = CsvToArray(UpdateFieldsCsv)
    TransferValidateFields Cn, TableName, fields, runInfo, "Source"
    TransferValidateFields POSConnection, TableName, fields, runInfo, "Destination"

    Set rsSchema = OpenSchemaRecordset(Cn, TableName)
    Set rsSrc = New ADODB.Recordset
    rsSrc.Open SourceSQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rsSrc.EOF Then
        sqlUpdate = BuildUpdateWithoutKeysSQLFromRS(TableName, rsSrc, fields, WhereClause, rsSchema)
        gTransferLastSQL = sqlUpdate
        SafeExecuteNonQuery POSConnection, sqlUpdate, runInfo, TableName & " SINGLE UPDATE"
        SyncSingleRowUpdate.Updated = 1
    Else
        SyncSingleRowUpdate.Skipped = 1
    End If

CleanExit:
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Exit Function

EH:
    SyncSingleRowUpdate.Failed = SyncSingleRowUpdate.Failed + 1
    TransferRememberError runInfo, "Table=" & TableName & " - " & Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR SyncSingleRowUpdate " & TableName & ": " & Err.Description
    SafeCloseRSLocal rsSchema
    SafeCloseRSLocal rsSrc
    Err.Raise Err.Number, IIf(Len(Err.Source) = 0, "ModTransferFixes.SyncSingleRowUpdate", Err.Source), Err.Description
End Function

Private Sub SafeExecuteNonQuery(ByVal cnExec As ADODB.Connection, ByVal SQLText As String, ByRef runInfo As TTransferRun, ByVal StepName As String)
    On Error GoTo EH
    Dim affected As Long

    TransferEnsureConnectionState cnExec, "Execution connection", runInfo, StepName
    TransferWriteLog runInfo, StepName & " SQL begin"
    cnExec.Errors.Clear
    cnExec.Execute SQLText, affected, adCmdText
    TransferWriteLog runInfo, StepName & " SQL success, affected=" & CStr(affected)
    Exit Sub

EH:
    TransferRememberError runInfo, StepName & " - " & Err.Description, SQLText
    TransferWriteLog runInfo, StepName & " SQL failed: " & Err.Description
    RaiseTransferAdoError cnExec, StepName, SQLText
End Sub

Private Sub RaiseTransferAdoError(ByVal cnExec As ADODB.Connection, ByVal StepName As String, ByVal SQLText As String)
    Dim s As String
    Dim i As Long

    s = StepName & vbCrLf & Err.Description
    If Not cnExec Is Nothing Then
        If cnExec.Errors.Count > 0 Then
        s = s & vbCrLf & vbCrLf & "ADO errors:"
            For i = 0 To cnExec.Errors.Count - 1
                s = s & vbCrLf & "- " & cnExec.Errors(i).Description & " (NativeError=" & CStr(cnExec.Errors(i).NativeError) & ")"
            Next i
        End If
    End If
    s = s & vbCrLf & vbCrLf & "SQL:" & vbCrLf & SQLText

    Err.Raise vbObjectError + 9200, "ModTransferFixes", s
End Sub

Private Function OpenSchemaRecordset(ByVal cnSrc As ADODB.Connection, ByVal TableName As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT TOP 0 * FROM " & TableName, cnSrc, adOpenStatic, adLockReadOnly, adCmdText
    Set OpenSchemaRecordset = rs
End Function

Private Sub TransferValidateFields(ByVal cnSrc As ADODB.Connection, ByVal TableName As String, ByRef fields() As String, ByRef runInfo As TTransferRun, ByVal sideName As String)
    On Error GoTo EH

    Dim rs As ADODB.Recordset
    Dim i As Long

    Set rs = OpenSchemaRecordset(cnSrc, TableName)
    For i = LBound(fields) To UBound(fields)
        If Not RecordsetHasField(rs, fields(i)) Then
            Err.Raise vbObjectError + 9300, , sideName & " table [" & TableName & "] is missing field [" & fields(i) & "]"
        End If
    Next i

    SafeCloseRSLocal rs
    Exit Sub

EH:
    TransferRememberError runInfo, Err.Description, "SELECT TOP 0 * FROM " & TableName
    SafeCloseRSLocal rs
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function BuildInsertSQLFromRS(ByVal TableName As String, ByVal rsData As ADODB.Recordset, ByRef fields() As String, ByVal rsSchema As ADODB.Recordset) As String
    Dim i As Long
    Dim sCols As String
    Dim sVals As String

    For i = LBound(fields) To UBound(fields)
        If sCols <> "" Then
            sCols = sCols & ","
            sVals = sVals & ","
        End If
        sCols = sCols & fields(i)
        sVals = sVals & SqlLiteralByField(rsData.fields(fields(i)).Value, rsSchema.fields(fields(i)).Type)
    Next i

    BuildInsertSQLFromRS = "INSERT INTO " & TableName & " (" & sCols & ") VALUES (" & sVals & ")"
End Function

Private Function BuildUpdateSQLFromRS(ByVal TableName As String, ByVal rsData As ADODB.Recordset, ByRef updateFields() As String, ByRef keyFields() As String, ByVal rsSchema As ADODB.Recordset) As String
    Dim i As Long
    Dim sSet As String

    For i = LBound(updateFields) To UBound(updateFields)
        If sSet <> "" Then sSet = sSet & ", "
        sSet = sSet & updateFields(i) & " = " & SqlLiteralByField(rsData.fields(updateFields(i)).Value, rsSchema.fields(updateFields(i)).Type)
    Next i

    BuildUpdateSQLFromRS = "UPDATE " & TableName & " SET " & sSet & " WHERE " & BuildWhereFromRecordset(rsData, keyFields)
End Function

Private Function BuildUpdateWithoutKeysSQLFromRS(ByVal TableName As String, ByVal rsData As ADODB.Recordset, ByRef updateFields() As String, ByVal WhereClause As String, ByVal rsSchema As ADODB.Recordset) As String
    Dim i As Long
    Dim sSet As String

    For i = LBound(updateFields) To UBound(updateFields)
        If sSet <> "" Then sSet = sSet & ", "
        sSet = sSet & updateFields(i) & " = " & SqlLiteralByField(rsData.fields(updateFields(i)).Value, rsSchema.fields(updateFields(i)).Type)
    Next i

    If Trim$(WhereClause) = "" Then
        BuildUpdateWithoutKeysSQLFromRS = "UPDATE " & TableName & " SET " & sSet
    Else
        BuildUpdateWithoutKeysSQLFromRS = "UPDATE " & TableName & " SET " & sSet & " WHERE " & WhereClause
    End If
End Function

Private Function BuildWhereFromRecordset(ByVal rsData As ADODB.Recordset, ByRef keyFields() As String) As String
    Dim i As Long
    Dim s As String

    For i = LBound(keyFields) To UBound(keyFields)
        If s <> "" Then s = s & " AND "
        s = s & keyFields(i) & " = " & SqlLiteralByField(rsData.fields(keyFields(i)).Value, rsData.fields(keyFields(i)).Type)
    Next i

    BuildWhereFromRecordset = s
End Function

Private Function BuildCompositeKeyFromRecordset(ByVal rsData As ADODB.Recordset, ByRef keyFields() As String) As String
    Dim i As Long
    Dim part As String
    Dim s As String

    For i = LBound(keyFields) To UBound(keyFields)
        part = NormalizeKeyValue(rsData.Fields(keyFields(i)).Value)
        s = s & CStr(Len(part)) & ":" & part & ";"
    Next i

    BuildCompositeKeyFromRecordset = s
End Function

Private Function NormalizeKeyValue(ByVal v As Variant) As String
    If IsNull(v) Then
        NormalizeKeyValue = "<NULL>"
    ElseIf IsDate(v) Then
        NormalizeKeyValue = Format$(CDate(v), "yyyy-mm-dd HH:nn:ss")
    ElseIf IsNumeric(v) Then
        NormalizeKeyValue = Replace$(Trim$(CStr(v)), ",", ".")
    Else
        NormalizeKeyValue = CStr(v)
    End If
End Function

Private Function LoadExistingKeys(ByVal cnDest As ADODB.Connection, ByVal TableName As String, ByRef keyFields() As String, ByRef runInfo As TTransferRun) As Collection
    On Error GoTo EH

    Dim rs As ADODB.Recordset
    Dim sqlKeys As String
    Dim loadedCount As Long
    Dim keyText As String

    sqlKeys = "SELECT " & JoinCsv(keyFields) & " FROM " & TableName
    gTransferLastSQL = sqlKeys
    TransferWriteLog runInfo, TableName & " loading destination keys"

    Set LoadExistingKeys = New Collection
    Set rs = New ADODB.Recordset
    rs.Open sqlKeys, cnDest, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rs.EOF
        keyText = BuildCompositeKeyFromRecordset(rs, keyFields)
        AddCollectionKey LoadExistingKeys, keyText
        loadedCount = loadedCount + 1
        rs.MoveNext
    Loop

    TransferWriteLog runInfo, TableName & " destination keys loaded: " & CStr(loadedCount)
    SafeCloseRSLocal rs
    Exit Function

EH:
    TransferRememberError runInfo, "LoadExistingKeys " & TableName & " - " & Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR LoadExistingKeys " & TableName & ": " & Err.Description
    SafeCloseRSLocal rs
    Err.Raise Err.Number, IIf(Len(Err.Source) = 0, "ModTransferFixes.LoadExistingKeys", Err.Source), Err.Description
End Function

Private Function JoinCsv(ByRef values() As String) As String
    Dim i As Long
    Dim s As String

    For i = LBound(values) To UBound(values)
        If s <> "" Then s = s & ","
        s = s & values(i)
    Next i

    JoinCsv = s
End Function

Private Sub AddCollectionKey(ByRef keys As Collection, ByVal KeyText As String)
    On Error Resume Next
    keys.Add KeyText, KeyText
    Err.Clear
    On Error GoTo 0
End Sub

Private Function CollectionContainsKey(ByRef keys As Collection, ByVal KeyText As String) As Boolean
    On Error GoTo EH
    Dim tmp As Variant

    tmp = keys.Item(KeyText)
    CollectionContainsKey = True
    Exit Function

EH:
    CollectionContainsKey = False
End Function

Private Sub QueueBatchStatement(ByRef BatchSQL As String, ByRef BatchCount As Long, ByVal SqlText As String)
    If Trim$(SqlText) = "" Then Exit Sub

    If Trim$(BatchSQL) = "" Then
        BatchSQL = SqlText
    Else
        BatchSQL = BatchSQL & vbCrLf & SqlText
    End If

    BatchCount = BatchCount + 1
End Sub

Private Sub FlushQueuedBatch(ByVal cnExec As ADODB.Connection, ByRef BatchSQL As String, ByRef BatchCount As Long, ByRef SuccessCounter As Long, ByRef runInfo As TTransferRun, ByVal StepName As String)
    If BatchCount <= 0 Or Trim$(BatchSQL) = "" Then Exit Sub

    gTransferLastSQL = BatchSQL
    SafeExecuteNonQuery cnExec, BatchSQL, runInfo, StepName & " BATCH(" & CStr(BatchCount) & ")"
    SuccessCounter = SuccessCounter + BatchCount
    BatchSQL = ""
    BatchCount = 0
End Sub

Private Sub TransferValidateLoopConnections(ByRef runInfo As TTransferRun, ByVal StepName As String, ByVal RowCount As Long)
    TransferEnsureConnectionState Cn, "Server connection", runInfo, StepName
    TransferEnsureConnectionState POSConnection, "Branch connection", runInfo, StepName
    TransferWriteLog runInfo, StepName & " progress rows=" & CStr(RowCount)
End Sub

Private Sub TransferEnsureConnectionState(ByVal cnCheck As ADODB.Connection, ByVal ConnectionName As String, ByRef runInfo As TTransferRun, ByVal StepName As String)
    On Error GoTo EH

    If cnCheck Is Nothing Then
        Err.Raise vbObjectError + 9401, , ConnectionName & " is not initialized"
    End If

    If cnCheck.State = adStateClosed Then
        Err.Raise vbObjectError + 9402, , ConnectionName & " is closed"
    End If

    Exit Sub

EH:
    TransferRememberError runInfo, StepName & " - " & Err.Description, gTransferLastSQL
    TransferWriteLog runInfo, "ERROR " & StepName & ": " & Err.Description
    Err.Raise Err.Number, IIf(Len(Err.Source) = 0, "ModTransferFixes.TransferEnsureConnectionState", Err.Source), Err.Description
End Sub

Private Function RecordExists(ByVal cnCheck As ADODB.Connection, ByVal SQLText As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open SQLText, cnCheck, adOpenForwardOnly, adLockReadOnly, adCmdText
    RecordExists = Not rs.EOF
    SafeCloseRSLocal rs
End Function

Private Function RecordsetHasField(ByVal rs As ADODB.Recordset, ByVal FieldName As String) As Boolean
    On Error GoTo EH
    Dim tmp As String
    tmp = rs.fields(FieldName).Name
    RecordsetHasField = True
    Exit Function
EH:
    RecordsetHasField = False
End Function

Private Function SqlLiteralByField(ByVal v As Variant, ByVal AdoType As DataTypeEnum) As String
    If IsNull(v) Then
        SqlLiteralByField = "NULL"
        Exit Function
    End If

    Select Case AdoType
        Case adTinyInt, adSmallInt, adInteger, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt
            SqlLiteralByField = Trim$(CStr(v))
        Case adSingle, adDouble, adCurrency, adDecimal, adNumeric
            SqlLiteralByField = Replace$(Trim$(CStr(v)), ",", ".")
        Case adBoolean
            If CBool(v) Then
                SqlLiteralByField = "1"
            Else
                SqlLiteralByField = "0"
            End If
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            SqlLiteralByField = "'" & Format$(CDate(v), "yyyy-mm-dd HH:nn:ss") & "'"
        Case Else
            SqlLiteralByField = "N'" & Replace$(CStr(v), "'", "''") & "'"
    End Select
End Function

Private Function CsvToArray(ByVal CsvText As String) As String()
    Dim arr() As String
    Dim raw() As String
    Dim i As Long

    raw = Split(CsvText, ",")
    ReDim arr(LBound(raw) To UBound(raw)) As String
    For i = LBound(raw) To UBound(raw)
        arr(i) = Trim$(raw(i))
    Next i
    CsvToArray = arr
End Function

Private Function JoinArray(ByRef arr1() As String, ByRef arr2() As String) As String()
    Dim total() As String
    Dim i As Long
    Dim p As Long

    ReDim total(LBound(arr1) To UBound(arr1) + UBound(arr2) + 1) As String

    p = LBound(total)
    For i = LBound(arr1) To UBound(arr1)
        total(p) = arr1(i)
        p = p + 1
    Next i
    For i = LBound(arr2) To UBound(arr2)
        total(p) = arr2(i)
        p = p + 1
    Next i

    JoinArray = total
End Function

Private Sub AccumulateCounters(ByRef cnt As TTransferCounters, ByRef TotalInserted As Long, ByRef TotalUpdated As Long, ByRef TotalFailed As Long, ByRef TotalSkipped As Long)
    TotalInserted = TotalInserted + cnt.Inserted
    TotalUpdated = TotalUpdated + cnt.Updated
    TotalFailed = TotalFailed + cnt.Failed
    TotalSkipped = TotalSkipped + cnt.Skipped
End Sub

Private Function TransferRunStart(ByVal Prefix As String) As TTransferRun
    Dim t As TTransferRun
    Dim p As String

    t.StartedAt = Now
    t.SessionCode = Prefix & "_" & Format$(Now, "yyyymmdd_hhnnss")
    p = App.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    t.LogFile = p & "TransferFix_" & t.SessionCode & ".log"

    TransferRunStart = t
End Function

Private Sub TransferWriteLog(ByRef runInfo As TTransferRun, ByVal Msg As String)
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open runInfo.LogFile For Append As #ff
    Print #ff, Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & Msg
    Close #ff
End Sub

Private Sub TransferRememberError(ByRef runInfo As TTransferRun, ByVal ErrText As String, ByVal SQLText As String)
    If Trim$(runInfo.FirstError) = "" Then
        runInfo.FirstError = ErrText
        runInfo.FirstErrorSQL = SQLText
    End If
End Sub

Private Function BuildTransferFailureMessage(ByRef runInfo As TTransferRun, ByVal PrefixText As String) As String
    Dim s As String

    s = PrefixText & vbCrLf
    If Trim$(runInfo.FirstError) <> "" Then
        s = s & "Reason: " & runInfo.FirstError & vbCrLf
    Else
        s = s & "Reason: " & Err.Description & vbCrLf
    End If

    If Trim$(runInfo.FirstErrorSQL) <> "" Then
        s = s & vbCrLf & "Last SQL:" & vbCrLf & runInfo.FirstErrorSQL & vbCrLf
    End If

    s = s & vbCrLf & "Trace file: " & runInfo.LogFile
    BuildTransferFailureMessage = s
End Function

Private Sub SafeCloseRSLocal(ByRef rs As ADODB.Recordset)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
End Sub
