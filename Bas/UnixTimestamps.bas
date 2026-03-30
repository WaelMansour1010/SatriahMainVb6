Attribute VB_Name = "UnixTimestamps"
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 63) As Byte
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 63) As Byte
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" ( _
    ByRef TimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function GetTimeZoneInformationForYear Lib "kernel32" ( _
    ByVal wYear As Integer, _
    ByVal pdtzi As Long, _
    ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" ( _
    ByRef TimeZoneInformation As TIME_ZONE_INFORMATION, _
    ByRef UniversalTime As SYSTEMTIME, _
    ByRef LocalTime As SYSTEMTIME) As Long
    
Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" ( _
    ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION, _
    ByRef lpLocalTime As SYSTEMTIME, _
    ByRef lpUniversalTime As SYSTEMTIME) As Long
    Public counterforitems As Double

Public Function OpenPostHttpRequest() As Boolean

 
End Function

Public Function WebRequest(URL As String) As String
' On Error GoTo errtrap
    Dim http As MSXML2.XMLHTTP
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
     http.Open "GET", URL, False, "testuser.ksa@gmail.com", "123456"
 http.setRequestHeader "Content-Type", "application/json"
 http.setRequestHeader "Transfer-Encoding", "chunked"
 http.setRequestHeader "Access-Control-Allow-Methods", "OPTIONS, GET, POST, PUT, DELETE"
 http.setRequestHeader "Access-Control-Allow-Origin", "*"
 http.setRequestHeader "Access-Control-Allow-Credentials", "true"
' http.setRequestHeader "Connection", " keep -alive"
 http.setRequestHeader "Access-Control-Allow-Headers", " accept, origin, x-requested-with, authorization, content-type"
 http.setRequestHeader "Server: nginx", "nginx"
 
 http.setRequestHeader "Authorization", "Token bb626c715b28385583fb372cd63bda26"
 http.send


  WebRequest = http.responseText
    Set http = Nothing
End Function
Public Function WebRequestPHP(URL As String, Optional GetBalance As Boolean) As String

    Dim DataToSend As String
    Dim objXML     As Object
    Dim Message    As String
    Dim authKey    As String
    Dim mobiles    As String
    Dim sender     As String
    Dim route      As String
 
    Set objXML = CreateObject("WinHttp.WinHttpRequest.5.1")
    objXML.Open "post", URL, False

    'objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'objXML.setRequestHeader "Content-Type", "text/html; charset=Windows-1256"
 
    'objXML.setRequestHeader "Content-Type", "text/html; charset=utf-8"

    If GetBalance = False Then
        'objXML.send "authkey=" + authKey + "&mobiles=" + mobiles + "&message=" + Message + "&sender=" + sender + "&route=" + route
        objXML.send
    Else

        objXML.send
    End If
    If Len(objXML.responseText) > 0 Then

        WebRequestPHP = objXML.responseText
        'MsgBox objXML.responseText
    End If

End Function
Function ChekInvoiceNoPurchasemanualExist(Transaction_ID As Double, CusID As Double, ManualNO As String, Optional ByRef NoteSerialT As String) As Boolean
 
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
 
sql = "select * from  Transactions where Transaction_Type=22 and  CusID=" & CusID & " and  Transaction_ID <>" & Transaction_ID & " and ManualNO='" & ManualNO & "'"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
ChekInvoiceNoPurchasemanualExist = True
NoteSerialT = IIf(IsNull(Rs5("NoteSerial1").value), "", Rs5("NoteSerial1").value)
Else
ChekInvoiceNoPurchasemanualExist = False
End If
End Function


Public Function CheckmanyAccount(Optional ByRef str As String = "") As Boolean

    Dim sql As String
    Dim rs As New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
sql = " SELECT     dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name "
Else
sql = " SELECT     dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_NameEng "
End If
sql = sql & " FROM         dbo.ACCOUNTS LEFT OUTER JOIN"
sql = sql & "                      dbo.TblUserAccount ON dbo.ACCOUNTS.Account_ID = dbo.TblUserAccount.Account_ID"
If user_id <> 1 Then
sql = sql & "    Where (dbo.TblUserAccount.UserID = " & user_id & ")"
Else
   CheckmanyAccount = False
   Exit Function
  End If
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
str = sql
        CheckmanyAccount = True
 
       
    Else
   CheckmanyAccount = False
    End If

End Function

Public Function GetProjectByUser() As String
Dim My_SQL As String
My_SQL = ""
            My_SQL = " and (projects.ID in (SELECT    TblProjectUser.ProjectID"
            My_SQL = My_SQL & " From dbo.TblProjectUser"
            My_SQL = My_SQL & "      where TblProjectUser.UserID=" & user_id & ")"
            My_SQL = My_SQL & " or  projects.ID not in (SELECT     TblProjectUser.ProjectID"
            My_SQL = My_SQL & " From dbo.TblProjectUser))"
      GetProjectByUser = My_SQL
End Function
Public Function GetExpensessPerstage(Optional Transaction_ID As Double, Optional Transaction_Type As Integer, Optional StoreId2 As Integer) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM((dbo.Transaction_Details.Price * dbo.Transaction_Details.showPrice) / (dbo.Transactions.Transaction_NetValue - ISNULL(dbo.Transactions.VAT, 0) "
sql = sql & "                     - ISNULL(dbo.Transactions.AddValue, 0))) AS Rate"
sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
sql = sql & " Where (dbo.Transaction_Details.Transaction_ID = " & Transaction_ID & ") And (dbo.transactions.Transaction_Type = " & Transaction_Type & ") And (dbo.Transaction_Details.StoreID2 = " & StoreId2 & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetExpensessPerstage = (IIf(IsNull(rs2("Rate").value), 0, rs2("Rate").value)) / 100
Else
GetExpensessPerstage = 0
End If
End Function
Public Function GetCheckHideAccount(Optional Typ As Integer = 0, Optional AccountCode As String) As Boolean
Dim My_SQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
            My_SQL = " SELECT    AccountCode"
            My_SQL = My_SQL & " From AccountSetting"
            My_SQL = My_SQL & " where  AccountCode='" & AccountCode & "'"
           If Typ = 0 Then
           My_SQL = My_SQL & " and     Entries = 1"
           ElseIf Typ = 1 Then
           My_SQL = My_SQL & " and     TrialBalance = 1"
           ElseIf Typ = 2 Then
           My_SQL = My_SQL & " and     TreeAccount = 1"
           End If
           
           rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          If rs2.RecordCount > 0 Then
          GetCheckHideAccount = True
          Else
          GetCheckHideAccount = False
      End If
End Function

Public Function GetHideAccount(Optional Typ As Integer = 0) As String
Dim My_SQL As String
My_SQL = ""
   
            My_SQL = " and (ACCOUNTS.Account_Code not in (SELECT    AccountCode"
           My_SQL = My_SQL & " From AccountSetting"
           If Typ = 0 Then
           My_SQL = My_SQL & " WHERE    Entries = 1))"
           ElseIf Typ = 1 Then
           My_SQL = My_SQL & " WHERE    TrialBalance = 1))"
            ElseIf Typ = 2 Then
           My_SQL = My_SQL & " WHERE    TreeAccount = 1))"
           End If
      GetHideAccount = My_SQL
End Function

Public Function get_transaction_NoteSerial1ByiDTemp(Transaction_ID As Double) As String

    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from Transactions where Transaction_ID=" & Transaction_ID
    'Sql = Sql & " and " & Transaction_Type
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        get_transaction_NoteSerial1ByiDTemp = ""
    Else
        get_transaction_NoteSerial1ByiDTemp = IIf(IsNull(rs("NoteSerial1").value), 0, rs("NoteSerial1").value)
    End If

End Function

Public Function ToLocal(ByVal UTCDateTime As Date) As Date
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION

    With stUTC
        .wYear = year(UTCDateTime)
        .wMonth = Month(UTCDateTime)
        .wDay = day(UTCDateTime)
        .wHour = Hour(UTCDateTime)
        .wMinute = Minute(UTCDateTime)
        .wSecond = Second(UTCDateTime)
        .wMilliseconds = 0
    End With

    'Requires Windows Vista or later:
    GetTimeZoneInformationForYear stUTC.wYear, 0, TimeZoneInfo
    'Fallback for pre-Longhorn Windows NT (e.g. Windows XP):
    'GetTimeZoneInformation TimeZoneInfo
    
    If SystemTimeToTzSpecificLocalTime(TimeZoneInfo, stUTC, stLocal) = 0 Then
        Err.Raise &H8004CC02, _
                  "ToLocal", _
                  "System error " & CStr(Err.LastDllError) _
                & "  calling SystemTimeToTzSpecificLocalTime"
    End If
    With stLocal
        ToLocal = DateSerial(.wYear, .wMonth, .wDay) _
                + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function ToUTC(ByVal LocalDateTime As Date) As Date
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    
    With stLocal
        .wYear = year(LocalDateTime)
        .wMonth = Month(LocalDateTime)
        .wDay = day(LocalDateTime)
        .wHour = Hour(LocalDateTime)
        .wMinute = Minute(LocalDateTime)
        .wSecond = Second(LocalDateTime)
        .wMilliseconds = 0
    End With
    Dim TimeZoneInfo As TIME_ZONE_INFORMATION
    
    'Requires Windows Vista or later:
    GetTimeZoneInformationForYear stLocal.wYear, 0, TimeZoneInfo
    'Fallback for pre-Longhorn Windows NT (e.g. Windows XP):
    'GetTimeZoneInformation TimeZoneInfo
    
    If TzSpecificLocalTimeToSystemTime(TimeZoneInfo, stLocal, stUTC) = 0 Then
        Err.Raise &H8004CC04, _
                  "ToUTC", _
                  "System error " & CStr(Err.LastDllError) _
                & "  calling TzSpecificLocalTimeToSystemTime"
    End If
    With stUTC
        ToUTC = DateSerial(.wYear, .wMonth, .wDay) _
              + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function UnixToDate(ByVal UnixTimestamp As Variant) As Date
    'Unix Timestamps are always UTC.
    Const MAX_LONG As Long = &H7FFFFFFF
    Const MAX_LONG_PLUS_1 As Currency = 2147483648@
    Dim Seconds As Variant
    Dim temp As Double

    If VarType(UnixTimestamp) <> vbDecimal Then
        Err.Raise 5
    Else
        Seconds = UnixTimestamp / CDec(1000)
        If Seconds > MAX_LONG Then
            temp = CDbl(Seconds - CDec(MAX_LONG_PLUS_1))
            UnixToDate = DateAdd("s", temp, #1/19/2038 3:14:08 AM#)
        Else
            temp = CDbl(Seconds)
            UnixToDate = DateAdd("s", temp, #1/1/1970#)
        End If
    End If
End Function

