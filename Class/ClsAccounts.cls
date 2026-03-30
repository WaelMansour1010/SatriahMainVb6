VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Public Function Repated(StrAccountSerial As String, _
                        Optional StrAccountCode As String = "") As Boolean
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT ACCOUNTS.* From ACCOUNTS " & "WHERE (((ACCOUNTS.Account_Serial)='" & StrAccountSerial & "'))"

    If StrAccountCode <> "" Then
        StrSQL = StrSQL + " AND (ACCOUNTS.Account_Code <>'" & StrAccountCode & "')"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Repated = True
    Else
        Repated = False
    End If

    rs.Close
    Set rs = Nothing
End Function
 

Public Function Get_Account_Serial(StrAccountCode As String) As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT ACCOUNTS.* From ACCOUNTS " & "WHERE (((ACCOUNTS.Account_Code)='" & StrAccountCode & "'))"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Get_Account_Serial = ""
    Else
        Get_Account_Serial = IIf(IsNull(rs("Account_Serial").value), "", rs("Account_Serial").value)
    End If

    rs.Close
    Set rs = Nothing
End Function



Public Function Get_Account_Name(StrAccountCode As String) As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT ACCOUNTS.* From ACCOUNTS " & "WHERE (((ACCOUNTS.Account_Code)='" & StrAccountCode & "'))"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        Get_Account_Name = ""
    Else
        Get_Account_Name = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
    End If

    rs.Close
    Set rs = Nothing
End Function

