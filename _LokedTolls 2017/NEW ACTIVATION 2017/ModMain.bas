Attribute VB_Name = "ModMain"
Public Cn As New ADODB.Connection
Public SysSQLServerType As Integer
Public SysSQLServerName As String
Public SysSQLServerTypeTechnical As String
Public StrAppRegPath As String
Public SysSQLServerDataBaseName As String
Public SysSQLServerUserId As String
Public SysSQLServerUserpassword As String
  
 Public Sub Main()
  StrAppRegPath = "bisegypt\SimpleAccounting"
SysSQLServerType = Val(GetSetting(StrAppRegPath, "ServerCon", "ServerType", 0)) '0 loca 1 not 2 rem
SysSQLServerName = GetSetting(StrAppRegPath, "ServerCon", "ServerName", "")
SysSQLServerTypeTechnical = GetSetting(StrAppRegPath, "ServerCon", "SysSQLServerTypeTechnical", "0")

SysSQLServerDataBaseName = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

     SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
    SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")

 

   Set Cn = New ADODB.Connection
    With Cn
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15


 

       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & SysSQLServerDataBaseName & _
        ";Data Source=" & SysSQLServerName & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & SysSQLServerDataBaseName & _
                    ";Data Source=" & SysSQLServerName & ";Port=1433"
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & SysSQLServerDataBaseName & ";Data Source=" & SysSQLServerName
                End If
          End If

'.Open
End With
Form1.Show

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
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrTemp = "#" & StrRes & "#"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrTemp = "'" & StrRes & "'"
        End If
        SQLDate = StrTemp
    End If
 End Function



