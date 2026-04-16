$path = 'F:\Source Code\SatriahMain\Bas\registry.bas'
$enc = [System.Text.Encoding]::Default
$lines = [System.Collections.Generic.List[string]]::new()
[System.IO.File]::ReadAllLines($path, $enc) | ForEach-Object { [void]$lines.Add($_) }

function FindLineStartsWith($list, [string]$value, [int]$start = 0) {
    for($i = $start; $i -lt $list.Count; $i++) {
        if($list[$i].StartsWith($value)) { return $i }
    }
    throw 'Line not found: ' + $value
}

function FindExact($list, [string]$value, [int]$start = 0) {
    for($i = $start; $i -lt $list.Count; $i++) {
        if($list[$i] -eq $value) { return $i }
    }
    throw 'Line not found: ' + $value
}

function ReplaceBlock($list, [int]$start, [int]$end, [string[]]$newLines) {
    $count = $end - $start + 1
    $list.RemoveRange($start, $count)
    for($i = $newLines.Count - 1; $i -ge 0; $i--) {
        $list.Insert($start, $newLines[$i])
    }
}

$helperStart = FindLineStartsWith $lines 'Private Function GetSalesCounterStartValue(' 0
$voucherIndex = FindLineStartsWith $lines 'Public Function Voucher_coding(' ($helperStart + 1)

$helperText = @'
Private Const TEMP_SALES_SERIAL_TRACE_ENABLED As Boolean = True

Private Sub TraceSalesSerialEvent(ByVal EventName As String, ByVal Detail As String)
    On Error Resume Next
    Dim FileNo As Integer
    Dim TraceLine As String

    TraceLine = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & EventName & " | " & Detail
    Debug.Print TraceLine

    If TEMP_SALES_SERIAL_TRACE_ENABLED = False Then Exit Sub

    FileNo = FreeFile
    Open App.Path & "\sales_serial_trace.log" For Append As #FileNo
    Print #FileNo, TraceLine
    Close #FileNo
End Sub

Private Function SalesSqlText(ByVal Value As String) As String
    SalesSqlText = "'" & Replace(Trim$(Value), "'", "''") & "'"
End Function

Private Function GetSalesCounterStartValue(ByVal start_at As Double) As Long
    GetSalesCounterStartValue = CLng(start_at) - 1
End Function

Private Function GetSalesCounterTranCount() As Long
    Dim RsTran As ADODB.Recordset

    On Error GoTo ExitHandler

    Set RsTran = New ADODB.Recordset
    RsTran.Open "SELECT @@TRANCOUNT AS TranCount", Cn, adOpenForwardOnly, adLockReadOnly
    If Not RsTran.EOF Then
        GetSalesCounterTranCount = val(RsTran!TranCount & "")
    End If

ExitHandler:
    On Error Resume Next
    If Not RsTran Is Nothing Then
        If RsTran.State = adStateOpen Then RsTran.Close
    End If
    Set RsTran = Nothing
End Function

Private Sub NormalizeSalesCounterKey(ByVal numbering_type As Integer, _
                                     ByVal Prefix As String, _
                                     ByVal date1 As Date, _
                                     ByVal StoreCoding As Double, _
                                     ByVal StoreID As Integer, _
                                     ByVal IsByUser As Boolean, _
                                     ByVal mUserId As Long, _
                                     ByVal mSerPosString As String, _
                                     ByRef KeyPrefix As String, _
                                     ByRef PeriodYear As Long, _
                                     ByRef PeriodMonth As Long, _
                                     ByRef KeyStoreID As Long, _
                                     ByRef KeyUserID As Long, _
                                     ByRef KeySerPos As Long)
    KeyPrefix = Trim$(Prefix)
    PeriodYear = 0
    PeriodMonth = 0
    KeyStoreID = 0
    KeyUserID = 0
    KeySerPos = 0

    If numbering_type = 2 Then
        PeriodYear = val(mId(Format$(date1, "dd/mm/yyyy"), 7, 4))
        PeriodMonth = val(mId(Format$(date1, "dd/mm/yyyy"), 4, 2))
    ElseIf numbering_type = 3 Then
        PeriodYear = val(mId(Format$(date1, "dd/mm/yyyy"), 7, 4))
    End If

    If (numbering_type = 2 Or numbering_type = 3) And StoreCoding = True And StoreID <> 0 Then
        KeyStoreID = StoreID
    End If

    If IsByUser Then
        KeyUserID = mUserId
    ElseIf numbering_type = 2 Then
        If StoreCoding <> True And SystemOptions.BranchDigit <= 1 Then
            KeySerPos = val(mSerPosString)
        End If
    End If
End Sub

Private Function BuildSalesCounterKeyWhere(ByVal Transaction_Type As Integer, _
                                           ByVal BranchID As Integer, _
                                           ByVal numbering_type As Integer, _
                                           ByVal KeyPrefix As String, _
                                           ByVal PeriodYear As Long, _
                                           ByVal PeriodMonth As Long, _
                                           ByVal KeyStoreID As Long, _
                                           ByVal KeyUserID As Long, _
                                           ByVal KeySerPos As Long) As String
    BuildSalesCounterKeyWhere = "TransactionType = " & Transaction_Type
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND BranchID = " & BranchID
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND NumberingType = " & numbering_type
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND Prefix = " & SalesSqlText(KeyPrefix)
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND PeriodYear = " & PeriodYear
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND PeriodMonth = " & PeriodMonth
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND StoreID = " & KeyStoreID
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND UserID = " & KeyUserID
    BuildSalesCounterKeyWhere = BuildSalesCounterKeyWhere & " AND SerPos = " & KeySerPos
End Function

Private Function GetSalesCounterValue(ByVal Transaction_Type As Integer, _
                                      ByVal BranchID As Integer, _
                                      ByVal numbering_type As Integer, _
                                      ByVal Prefix As String, _
                                      ByVal date1 As Date, _
                                      ByVal StoreCoding As Double, _
                                      ByVal StoreID As Integer, _
                                      ByVal IsByUser As Boolean, _
                                      ByVal mUserId As Long, _
                                      ByVal mSerPosString As String, _
                                      ByVal start_at As Double, _
                                      ByVal AllocateNow As Boolean, _
                                      ByRef NextCounterValue As Long, _
                                      ByRef KeyDebug As String) As Boolean
    Dim RsCounter   As ADODB.Recordset
    Dim sqlCounter  As String
    Dim sqlWhere    As String
    Dim KeyPrefix   As String
    Dim PeriodYear  As Long
    Dim PeriodMonth As Long
    Dim KeyStoreID  As Long
    Dim KeyUserID   As Long
    Dim KeySerPos   As Long
    Dim seedValue   As Long

    On Error GoTo ErrHandler

    NormalizeSalesCounterKey numbering_type, Prefix, date1, StoreCoding, StoreID, IsByUser, mUserId, mSerPosString, KeyPrefix, PeriodYear, PeriodMonth, KeyStoreID, KeyUserID, KeySerPos
    sqlWhere = BuildSalesCounterKeyWhere(Transaction_Type, BranchID, numbering_type, KeyPrefix, PeriodYear, PeriodMonth, KeyStoreID, KeyUserID, KeySerPos)
    KeyDebug = "TransactionType=" & Transaction_Type & " | BranchID=" & BranchID & " | NumberingType=" & numbering_type & " | Prefix=" & IIf(KeyPrefix = "", "<NULL>", KeyPrefix) & " | PeriodYear=" & PeriodYear & " | PeriodMonth=" & PeriodMonth & " | StoreID=" & KeyStoreID & " | UserID=" & KeyUserID & " | SerPos=" & KeySerPos
    seedValue = GetSalesCounterStartValue(start_at)

    Set RsCounter = New ADODB.Recordset

    If AllocateNow Then
        TraceSalesSerialEvent "real allocate call", KeyDebug

        sqlCounter = "SELECT CounterValue FROM SerialCounters WITH (UPDLOCK, HOLDLOCK) WHERE " & sqlWhere
        RsCounter.Open sqlCounter, Cn, adOpenKeyset, adLockOptimistic
        If RsCounter.EOF Then
            RsCounter.Close
            sqlCounter = "INSERT INTO SerialCounters (TransactionType, BranchID, NumberingType, Prefix, PeriodYear, PeriodMonth, StoreID, UserID, SerPos, CounterValue, LastUpdated) VALUES ("
            sqlCounter = sqlCounter & Transaction_Type & ", " & BranchID & ", " & numbering_type & ", " & SalesSqlText(KeyPrefix) & ", " & PeriodYear & ", " & PeriodMonth & ", " & KeyStoreID & ", " & KeyUserID & ", " & KeySerPos & ", " & seedValue & ", GETDATE())"
            Cn.Execute sqlCounter
        Else
            RsCounter.Close
        End If

        sqlCounter = "UPDATE SerialCounters SET CounterValue = CounterValue + 1, LastUpdated = GETDATE() WHERE " & sqlWhere
        Cn.Execute sqlCounter

        sqlCounter = "SELECT CounterValue FROM SerialCounters WHERE " & sqlWhere
        RsCounter.Open sqlCounter, Cn, adOpenForwardOnly, adLockReadOnly
        If RsCounter.EOF Then GoTo ErrHandler
        NextCounterValue = CLng(RsCounter!CounterValue)
        TraceSalesSerialEvent "allocated counter value", KeyDebug & " | CounterValue=" & NextCounterValue
    Else
        TraceSalesSerialEvent "peek/validate call", KeyDebug

        sqlCounter = "SELECT CounterValue FROM SerialCounters WHERE " & sqlWhere
        RsCounter.Open sqlCounter, Cn, adOpenForwardOnly, adLockReadOnly
        If RsCounter.EOF Or IsNull(RsCounter!CounterValue) Then
            NextCounterValue = seedValue + 1
        Else
            NextCounterValue = CLng(RsCounter!CounterValue) + 1
        End If
        TraceSalesSerialEvent "peek counter value", KeyDebug & " | CounterValue=" & NextCounterValue
    End If

    GetSalesCounterValue = True

ExitHandler:
    On Error Resume Next
    If Not RsCounter Is Nothing Then
        If RsCounter.State = adStateOpen Then RsCounter.Close
    End If
    Set RsCounter = Nothing
    Exit Function

ErrHandler:
    TraceSalesSerialEvent "sales counter error", KeyDebug & " | Err=" & Err.Number & " | " & Err.Description
    GetSalesCounterValue = False
    Resume ExitHandler
End Function

Private Function BuildSalesVisibleSerial(ByVal IsByUser As Boolean, _
                                         ByVal numbering_type As Integer, _
                                         ByVal auto_sanad_no As String, _
                                         ByVal brancHcode As String, _
                                         ByVal storecode As String, _
                                         ByVal StoreCoding As Double, _
                                         ByVal StoreID As Integer, _
                                         ByVal mUserIdSerial As String, _
                                         ByVal mSerPosString As String) As String
    If numbering_type = 1 Then
        BuildSalesVisibleSerial = auto_sanad_no
        Exit Function
    End If

    If IsByUser Then
        If StoreCoding = True And StoreID <> 0 Then
            BuildSalesVisibleSerial = "1" & mUserIdSerial & brancHcode & storecode & auto_sanad_no
        Else
            BuildSalesVisibleSerial = "1" & mUserIdSerial & brancHcode & auto_sanad_no
        End If
    Else
        If StoreCoding = True And StoreID <> 0 Then
            BuildSalesVisibleSerial = brancHcode & storecode & auto_sanad_no
        ElseIf mSerPosString <> "" Then
            BuildSalesVisibleSerial = mSerPosString & brancHcode & auto_sanad_no
        Else
            BuildSalesVisibleSerial = brancHcode & auto_sanad_no
        End If
    End If
End Function

Private Function SalesSerialExists(ByVal Transaction_Type As Integer, _
                                   ByVal BranchID As Integer, _
                                   ByVal numbering_type As Integer, _
                                   ByVal Prefix As String, _
                                   ByVal date1 As Date, _
                                   ByVal StoreCoding As Double, _
                                   ByVal StoreID As Integer, _
                                   ByVal IsByUser As Boolean, _
                                   ByVal mUserId As Long, _
                                   ByVal mSerPosString As String, _
                                   ByVal VisibleSerial As String) As Boolean
    Dim RsDup       As ADODB.Recordset
    Dim sqlDup      As String
    Dim KeyPrefix   As String
    Dim PeriodYear  As Long
    Dim PeriodMonth As Long
    Dim KeyStoreID  As Long
    Dim KeyUserID   As Long
    Dim KeySerPos   As Long

    On Error GoTo ExitHandler

    NormalizeSalesCounterKey numbering_type, Prefix, date1, StoreCoding, StoreID, IsByUser, mUserId, mSerPosString, KeyPrefix, PeriodYear, PeriodMonth, KeyStoreID, KeyUserID, KeySerPos

    sqlDup = "SELECT COUNT(*) AS DuplicateCount FROM Transactions WHERE BranchId = " & BranchID & " AND Transaction_Type = " & Transaction_Type
    sqlDup = sqlDup & " AND NoteSerial1 = " & SalesSqlText(VisibleSerial)

    If KeyPrefix = "" Then
        sqlDup = sqlDup & " AND Prefix IS NULL"
    Else
        sqlDup = sqlDup & " AND Prefix = " & SalesSqlText(KeyPrefix)
    End If

    If numbering_type = 2 Then
        sqlDup = sqlDup & " AND YEAR(Transaction_Date) = " & PeriodYear & " AND MONTH(Transaction_Date) = " & PeriodMonth
    ElseIf numbering_type = 3 Then
        sqlDup = sqlDup & " AND YEAR(Transaction_Date) = " & PeriodYear
    End If

    If KeyStoreID <> 0 Then
        sqlDup = sqlDup & " AND StoreID = " & KeyStoreID
    End If

    If KeyUserID <> 0 Then
        sqlDup = sqlDup & " AND UserID = " & KeyUserID
    End If

    If KeySerPos <> 0 Or (IsByUser = False And numbering_type = 2 And StoreCoding <> True And SystemOptions.BranchDigit <= 1) Then
        sqlDup = sqlDup & " AND ISNULL(SerPos,0) = " & KeySerPos
    End If

    Set RsDup = New ADODB.Recordset
    RsDup.Open sqlDup, Cn, adOpenForwardOnly, adLockReadOnly
    If Not RsDup.EOF Then
        SalesSerialExists = (val(RsDup!DuplicateCount & "") > 0)
    End If

ExitHandler:
    On Error Resume Next
    If Not RsDup Is Nothing Then
        If RsDup.State = adStateOpen Then RsDup.Close
    End If
    Set RsDup = Nothing
End Function
'@

$helperLines = $helperText -split "`r?`n"
ReplaceBlock $lines $helperStart ($voucherIndex - 1) $helperLines

$voucherIndex = FindLineStartsWith $lines 'Public Function Voucher_coding(' 0
$declIndex = FindExact $lines '    Dim mSalesCounterValue As Long' $voucherIndex
if ($lines[$declIndex + 1] -ne '    Dim mAllocateSalesCounter As Boolean') {
    $lines.Insert($declIndex + 1, '    Dim mAllocateSalesCounter As Boolean')
    $lines.Insert($declIndex + 2, '    Dim mSalesVisibleSerial As String')
    $lines.Insert($declIndex + 3, '    Dim mSalesTraceKey    As String')
}

$startBlock1 = FindExact $lines '    If Transaction_Type = 21 And numbering_type <> 0 Then' $voucherIndex
$endBlock1 = $startBlock1
while($lines[$endBlock1 + 1] -ne '    If numbering_type = 1 Then '' الي') { $endBlock1++ }
$block1Text = @'
    If Transaction_Type = 21 And numbering_type <> 0 Then
        Askcount = noOfDigit
        If Askcount = 0 Then Askcount = 3

        mAllocateSalesCounter = (GetSalesCounterTranCount() > 0)
        If GetSalesCounterValue(Transaction_Type, my_branch, numbering_type, Prefix, date1, StoreCoding, StoreID, False, 0, mSerPosString, start_at, mAllocateSalesCounter, mSalesCounterValue, mSalesTraceKey) = False Then
            Voucher_coding = "error"
            Exit Function
        End If

        mSerInv = mSalesCounterValue

        If end_at <> 0 And mSalesCounterValue > end_at Then
            TraceSalesSerialEvent "sales counter limit", mSalesTraceKey & " | CounterValue=" & mSalesCounterValue & " | EndAt=" & end_at
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

        mSalesVisibleSerial = BuildSalesVisibleSerial(False, numbering_type, auto_sanad_no, brancHcode, zeropadding(storecode, Int(SystemOptions.StoreDigit)), StoreCoding, StoreID, "", mSerPosString)
        If SalesSerialExists(Transaction_Type, my_branch, numbering_type, Prefix, date1, StoreCoding, StoreID, False, 0, mSerPosString, mSalesVisibleSerial) Then
            TraceSalesSerialEvent "duplicate sales serial detected", mSalesTraceKey & " | VisibleSerial=" & mSalesVisibleSerial & " | AllocateNow=" & mAllocateSalesCounter
            Voucher_coding = "error"
            Exit Function
        End If

        GoTo BuildSalesVoucherCode
    End If
'@
$block1Lines = $block1Text -split "`r?`n"
ReplaceBlock $lines $startBlock1 $endBlock1 $block1Lines

$byUserIndex = FindLineStartsWith $lines 'Public Function Voucher_codingByUser(' ($voucherIndex + 1)
$declIndex2 = FindExact $lines '    Dim mSalesCounterValue As Long' $byUserIndex
if ($lines[$declIndex2 + 1] -ne '    Dim mAllocateSalesCounter As Boolean') {
    $lines.Insert($declIndex2 + 1, '    Dim mAllocateSalesCounter As Boolean')
    $lines.Insert($declIndex2 + 2, '    Dim mSalesVisibleSerial As String')
    $lines.Insert($declIndex2 + 3, '    Dim mSalesTraceKey    As String')
}

$startBlock2 = FindExact $lines '    If Transaction_Type = 21 And numbering_type <> 0 Then' $byUserIndex
$endBlock2 = $startBlock2
while($lines[$endBlock2 + 1] -ne '    If numbering_type = 1 Then '' الي') { $endBlock2++ }
$block2Text = @'
    If Transaction_Type = 21 And numbering_type <> 0 Then
        Askcount = noOfDigit
        If Askcount = 0 Then Askcount = 3

        mAllocateSalesCounter = (GetSalesCounterTranCount() > 0)
        If GetSalesCounterValue(Transaction_Type, my_branch, numbering_type, Prefix, date1, StoreCoding, StoreID, True, mUserId, mSerPosString, start_at, mAllocateSalesCounter, mSalesCounterValue, mSalesTraceKey) = False Then
            Voucher_codingByUser = "error"
            Exit Function
        End If

        If end_at <> 0 And mSalesCounterValue > end_at Then
            TraceSalesSerialEvent "sales counter limit", mSalesTraceKey & " | CounterValue=" & mSalesCounterValue & " | EndAt=" & end_at
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

        mSalesVisibleSerial = BuildSalesVisibleSerial(True, numbering_type, auto_sanad_no, zeropadding(CStr(my_branch), Int(SystemOptions.BranchDigit)), zeropadding(storecode, Int(SystemOptions.StoreDigit)), StoreCoding, StoreID, mUserIdSerial, mSerPosString)
        If SalesSerialExists(Transaction_Type, my_branch, numbering_type, Prefix, date1, StoreCoding, StoreID, True, mUserId, mSerPosString, mSalesVisibleSerial) Then
            TraceSalesSerialEvent "duplicate sales serial detected", mSalesTraceKey & " | VisibleSerial=" & mSalesVisibleSerial & " | AllocateNow=" & mAllocateSalesCounter
            Voucher_codingByUser = "error"
            Exit Function
        End If

        GoTo BuildSalesVoucherCodeByUser
    End If
'@
$block2Lines = $block2Text -split "`r?`n"
ReplaceBlock $lines $startBlock2 $endBlock2 $block2Lines

[System.IO.File]::WriteAllLines($path, $lines, $enc)
