Attribute VB_Name = "MyFunctions"
Public Enum TransNaV
    'Nav buttons index
    EnFirstTrans = 1
    EnLastTrans = 2
    EnNextTrans = 3
    EnPervTrans = 0
    EnSrchTrans = 4
End Enum
Public LastIdentityInsertTable As String
 Public Sub CloseRS(mrs)
 If mrs.State = adStateOpen Then
        If Not (mrs.EOF Or mrs.BOF) Then
            If mrs.EditMode <> adEditNone Then
                mrs.CancelUpdate
            End If
        End If

        mrs.Close
        
    End If
End Sub
Public Function GetUserSign(Optional ByVal mTransaction_ID As Long = 0, Optional ByVal ScreenName As String = "", Optional ByVal mLevel As Long = 0) As String

    Dim StrSQL  As String
    StrSQL = "(SELECT     TOP 1 UserSign "
    
    StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
    StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & Val(mTransaction_ID) & ") AND (dbo.ApprovalData.ScreenName = N'" & ScreenName & "') and ApprovalData.levelo = " & mLevel & ")"
    GetUserSign = StrSQL
End Function
Public Sub UpdateFiles(ByVal POSlServer As String, _
                       ByVal POSDb As String, _
                       ByVal mTableName As String, _
                       Optional ByVal mFieldName As String = "Id", _
                       Optional ByVal mWhere As String = "")
    
    Dim NoOFItem_POS    As Double
    Dim NoOFItem_Server As Double
   
    Dim Rs3             As New ADODB.Recordset
    Dim MaxItem_POS     As Double
    Dim MaxItem_Server  As Double
    Dim ss              As String
    
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    If mTableName <> "TblOptions55" Then
    sql = " select count (" & mFieldName & " ) As NoOfitems ,max(" & mFieldName & " ) as MaxItemid from " & mTableName
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
    Rs3.Close
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    ' MsgBox "Step 1"
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close
    End If
    ' MsgBox "Item Server" & NoOFItem_Server
    ' MsgBox "Item Pos" & NoOFItem_POS
    'step 2
    ' Exit Sub
             
    s = ""
             
    Dim mPosD    As String
    Dim mServerD As String
    mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
    mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
    
    
    
        ' ≈⁄œ«œ „ €Ì—«  «·—»ÿ
  '  Dim mPosD As String, mServerD As String
    ' ﬁ«⁄œ… »Ì«‰«  «·‰ﬁÿ…:
   ' mPosD = "[" & POSlServer.text & "]." & POSDb & ".dbo."
    ' ﬁ«⁄œ… »Ì«‰«  «·”Ì—›— «·»⁄Ìœ:
    mServerD = "[RemoteServer10]." & ServerDb & ".dbo."

    ' ≈⁄œ«œ « ’«· «·‰ﬁÿ… (POSConnection)
'    Set POSConnection = New ADODB.Connection
'    With POSConnection
'        .CommandTimeout = 5000
'        .CursorLocation = adUseClient
'        .ConnectionTimeout = 5000
'        If POSServer = "" Then POSServer = POSlServer
'        If SysSQLServerType = 1 Then
'            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                ";Initial Catalog=" & POSDb & ";Data Source=" & POSServer
'        ElseIf SysSQLServerType = 2 Then
'             If SysSQLServerTypeTechnical = "0" Then
'                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & POSDb & _
'                                     ";Data Source=" & POSlServer & ";Port=1433"
'              Else
'                 .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
'                                     ";Initial Catalog=" & POSDb & ";Data Source=" & POSlServer
'            End If
'        End If
'        .Open
'    End With


    'mServerD = ServerDb & ".dbo."
         
    ' Text4 = s
    ' Exit Sub
    mPosD = ""
   ' POSConnection.Execute "SET IDENTITY_INSERT " & mTableName & " ON"
     If mTableName <> "TblLink_Item_To_Store_Details3" And mTableName <> "TblLink_Item_To_Store_Details1" Then
        SetIdentityInsertSafe POSConnection, mTableName, POSDb
    End If

    s = GetSql(mServerD, mPosD, mTableName, mFieldName)
    FRMTRansferData3.Text4 = s
    POSConnection.Execute s
    ' MsgBox "Step 4"
            
    '  MsgBox " „ ‰ﬁ· »Ì«‰«  «·»Ì«‰« "
    '  cmdUdateFiles.Enabled = False

End Sub


Public Sub SetIdentityInsertSafe(ByRef Conn As ADODB.Connection, ByVal tablename As String, ByVal DbName As String)
    On Error Resume Next

    ' ≈Ìﬁ«› «·ÃœÊ· «·”«»ﬁ ≈‰ ÊÃœ
  On Error GoTo errHandler
    
    Dim rs As New ADODB.Recordset
    Dim sql As String

    ' «· Õﬁﬁ Â· «·ÃœÊ· ÌÕ ÊÌ ⁄·Ï ⁄„Êœ Identity
    sql = "SELECT COLUMN_NAME FROM " & DbName & ".INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & tablename & "' AND COLUMNPROPERTY(OBJECT_ID('" & DbName & ".dbo." & tablename & "'), COLUMN_NAME, 'IsIdentity') = 1"
    
    rs.Open sql, Conn, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        ' «·ÃœÊ· ·« ÌÕ ÊÌ ⁄·Ï identity ? ·«  ›⁄· ‘Ì¡
        Exit Sub
    End If
    rs.Close

    ' ≈Ìﬁ«› «·ÃœÊ· «·”«»ﬁ ≈‰ ÊÃœ
    If LastIdentityInsertTable <> "" And LastIdentityInsertTable <> tablename Then
        Conn.Execute "SET IDENTITY_INSERT " & DbName & ".dbo." & LastIdentityInsertTable & " OFF"
    End If

    '  ‘€Ì· IDENTITY_INSERT ··ÃœÊ· «·ÃœÌœ
    Conn.Execute "SET IDENTITY_INSERT " & DbName & ".dbo." & tablename & " ON"
    LastIdentityInsertTable = tablename
    Exit Sub

errHandler:
    MsgBox "SetIdentityInsertSafe Error: " & Err.Description
    Err.Clear
End Sub

Public Sub UpdateFilesFromPos(ByVal POSlServer As String, _
                              ByVal POSDb As String, _
                              ByVal mTableName As String, _
                              Optional ByVal mFieldName As String = "Id", _
                              Optional ByVal mWhere As String = "", _
                              Optional ByVal mPOSlServer As String = "")
    
    Dim NoOFItem_POS    As Double
    Dim NoOFItem_Server As Double
   
    Dim Rs3             As New ADODB.Recordset
    Dim MaxItem_POS     As Double
    Dim MaxItem_Server  As Double
    Dim ss              As String
    
    ss = "     USE " & ServerDb & vbNewLine
    
    Cn.Execute ss
    ss = "USE " & POSDb & vbNewLine
    POSConnection.Execute ss
    
    sql = " select count (" & mFieldName & " ) As NoOfitems ,max(" & mFieldName & " ) as MaxItemid from " & mTableName
     
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
    If Rs3.RecordCount > 0 Then
        NoOFItem_POS = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_POS = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
   
    End If
    Rs3.Close
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    ' MsgBox "Step 1"
    If Rs3.RecordCount > 0 Then
        NoOFItem_Server = IIf(IsNull(Rs3("NoOfitems").Value), 0, Rs3("NoOfitems").Value)
        MaxItem_Server = IIf(IsNull(Rs3("MaxItemid").Value), 0, Rs3("MaxItemid").Value)
    End If
    Rs3.Close
    
    ' MsgBox "Item Server" & NoOFItem_Server
    ' MsgBox "Item Pos" & NoOFItem_POS
    'step 2
    ' Exit Sub

    'MsgBox "Step 3"
    Dim s As String
             
    s = ""
             
    Dim mPosD    As String
    Dim mServerD As String
    mPosD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
    '   mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
    If mPOSlServer <> "" Then
        mServerD = mPOSlServer
    Else
        mServerD = "[" & mPOSlServer & "]." & POSDb & ".dbo."
    End If
         
    ' Text4 = s
    ' Exit Sub
    '  mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
          
    s = "Select * from " & mTableName & " Where cusid = -5"

    s = GetSql2(mServerD, mPosD, mTableName, mFieldName, mWhere)
    Cn.Execute s
    ' MsgBox "Step 4"
            
    '  MsgBox " „ ‰ﬁ· »Ì«‰«  «·»Ì«‰« "
    '  cmdUdateFiles.Enabled = False

End Sub

Public Function GetSql2(ByVal POSlServer As String, _
                        ByVal POSDb As String, _
                        ByVal mTableName As String, _
                        Optional ByVal mFieldName As String = "Id", _
                        Optional ByVal mWhere As String = "") As String
    
    Dim s As String
    s = "Select * from " & mTableName & " Where 1 = -1  " ' & mWhere
    Dim rs As New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim mName  As String
    Dim mName2 As String
    Dim ii     As Long
    mName = ""
    mName2 = ""
    ii = 0
    For ii = 0 To rs.Fields.Count - 1
        If mName = "" Then
            If mTableName = "Transaction_Details" Or mTableName = "TransactionValueAdded" Then
                If UCase(rs(ii).Name) = "ID" Then
                    GoTo NextCol
                End If
            End If
            mName = mName & rs(ii).Name
            mName2 = mName2 & "TT." & Trim(rs(ii).Name)
        Else
            If rs(ii).Name <> "MainOperationID" Then
                If mTableName = "Transaction_Details" Or mTableName = "TransactionValueAdded" Then
                    If UCase(rs(ii).Name) = "ID" Then
                        GoTo NextCol
                    End If
                End If
                If ii = 7 Or ii = 14 Or ii = 21 Or ii = 28 Or ii = 35 Or ii = 42 Or ii = 49 Or ii = 56 Or ii = 66 Or ii = 77 Or ii = 82 Or ii = 90 Or ii = 95 Or ii = 100 Or ii = 107 Or ii = 113 Or ii = 117 Or ii = 125 Or ii = 130 Or ii = 140 Or ii = 149 Or ii = 156 Or ii = 162 Or ii = 170 Or ii = 177 Or ii = 182 Or ii = 190 Or ii = 195 Or ii = 202 Or ii = 210 Or ii = 215 Or ii = 221 Or ii = 230 Or ii = 238 Or ii = 244 Or ii = 250 Or ii = 257 Or ii = 262 Or ii = 270 Or ii = 277 Or ii = 282 Or ii = 290 Or ii = 297 Or ii = 305 Or ii = 312 Or ii = 320 Or ii = 328 Or ii = 335 Or ii = 342 Then
                    mName = mName & "," & rs(ii).Name & vbNewLine
                    mName2 = mName2 & ",TT." & rs(ii).Name & vbNewLine
                Else
                    mName = mName & "," & rs(ii).Name
                    mName2 = mName2 & ",TT." & rs(ii).Name

                End If
            End If
        End If
NextCol:
        ' ii = ii + 1
    Next

    s = " INSERT INTO " & POSDb & mTableName & " (" & mName & ")"
    s = s & " SELECT  " & mName2
    s = s & " FROM   " & POSlServer & "" & mTableName & " TT"
    s = s & " WHERE  1 = 1 "
    'T2." & mFieldName & "   NOT IN (SELECT " & mFieldName & " "
    's = s & "                                      FROM   " & POSDb & mTableName & " );"
    If mWhere <> "" Then
        s = s & "  And " & mWhere
    End If
    
    GetSql2 = s
    
End Function



Public Function GetSql(ByVal POSlServer As String, _
                       ByVal POSDb As String, _
                       ByVal mTableName As String, _
                       Optional ByVal mFieldName As String = "Id", _
                       Optional ByVal mWhere As String = "") As String

    Dim s As String
    Dim rs As New ADODB.Recordset
    Dim mName As String, mName2 As String, updates As String
    Dim ii As Long

    s = "SELECT * FROM " & mTableName & " WHERE 1 = -1"
    rs.Open s, Cn, adOpenStatic, adLockReadOnly

    mName = ""
    mName2 = ""
    updates = ""

    For ii = 0 To rs.Fields.Count - 1
        If rs.Fields(ii).Name <> "MainOperationID" Then
            If mName = "" Then
                If (mTableName = "TblLink_Item_To_Store_Details1" Or mTableName = "TblLink_Item_To_Store_Details2" Or mTableName = "TblLink_Item_To_Store_Details3") And UCase(Trim(rs.Fields(ii).Name) & "") = "ID" Then
                Else
                    mName = rs.Fields(ii).Name
                    mName2 = "T2." & rs.Fields(ii).Name
                End If
                
            Else
                mName = mName & "," & rs.Fields(ii).Name
                mName2 = mName2 & ",T2." & rs.Fields(ii).Name
            End If

            If UCase(rs.Fields(ii).Name) <> UCase(mFieldName) Then
                If updates = "" Then
                    updates = "T1." & rs.Fields(ii).Name & " = T2." & rs.Fields(ii).Name
                Else
                    updates = updates & "," & vbCrLf & "T1." & rs.Fields(ii).Name & " = T2." & rs.Fields(ii).Name
                End If
            End If
        End If
    Next

    ' INSERT statement
    s = "INSERT INTO " & POSDb & mTableName & " (" & mName & ") " & vbCrLf
    s = s & "SELECT " & mName2 & vbCrLf
    s = s & "FROM " & POSlServer & mTableName & " T2 " & vbCrLf
    s = s & "WHERE T2." & mFieldName & " NOT IN (SELECT " & mFieldName & " FROM " & POSDb & mTableName & ")"

    If mWhere <> "" Then
        s = s & " AND (" & mWhere & ")"
    End If

    s = s & ";" & vbCrLf & vbCrLf

    ' UPDATE statement
    s = s & "UPDATE T1 SET " & vbCrLf & updates & vbCrLf
    s = s & "FROM " & POSDb & mTableName & " T1 " & vbCrLf
    s = s & "JOIN " & POSlServer & mTableName & " T2 ON T1." & mFieldName & " = T2." & mFieldName

    GetSql = s
End Function

Public Function GetSql_Old(ByVal POSlServer As String, _
                       ByVal POSDb As String, _
                       ByVal mTableName As String, _
                       Optional ByVal mFieldName As String = "Id", _
                       Optional ByVal mWhere As String = "") As String
    Dim s As String
    s = "Select * from " & mTableName & " Where 1 = -1  " ' & mWhere
    Dim rs As New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim mName  As String
    Dim mName2 As String
    Dim ii     As Long
    mName = ""
    mName2 = ""
    For ii = 0 To rs.Fields.Count - 1
        If mName = "" Then
            mName = mName & rs(ii).Name
            mName2 = mName2 & "T2." & Trim(rs(ii).Name)
        Else
            If rs(ii).Name <> "MainOperationID" Then
                mName = mName & "," & rs(ii).Name & vbNewLine
                mName2 = mName2 & ",T2." & rs(ii).Name & vbNewLine
            End If
        End If
        ii = ii + 1
    Next

    s = " INSERT INTO " & POSDb & mTableName & " (" & mName & ")"
    s = s & " SELECT " & mName2
    s = s & " FROM   " & POSlServer & "" & mTableName & " T2"
    s = s & " WHERE  T2." & mFieldName & "   NOT IN (SELECT " & mFieldName & " "
    s = s & "                                      FROM   " & POSDb & mTableName & " );"
    If mWhere <> "" Then
        s = s & "  And mWhere                                     "
    End If
    
    GetSql_Old = s
End Function
'samy to fix windows 10 error
Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
Set WshShell = Nothing
End Sub





Public Sub LogErrDetailed(ByVal ProcName As String, ByVal ex As ErrObject, _
                          Optional ByVal ErlLine As Long = 0, _
                          Optional ByVal ExtraInfo As String = "", _
                          Optional ByVal Cn As ADODB.Connection = Nothing, _
                          Optional ByVal POSCn As ADODB.Connection = Nothing, _
                          Optional ByVal LastSQL As String = "")

    Dim Msg As String
    Msg = "? Error in: " & ProcName & vbCrLf & _
          "Err.Number: " & ex.Number & vbCrLf & _
          "Err.Description: " & ex.Description & vbCrLf & _
          "Err.Source: " & ex.Source & vbCrLf & _
          "Erl (Line): " & ErlLine & vbCrLf

    If Len(ExtraInfo) > 0 Then
        Msg = Msg & vbCrLf & "Extra:" & vbCrLf & ExtraInfo & vbCrLf
    End If

    If Len(LastSQL) > 0 Then
        Msg = Msg & vbCrLf & "Last SQL:" & vbCrLf & Left$(LastSQL, 4000) & vbCrLf
    End If

    Msg = Msg & vbCrLf & "---- Connections ----" & vbCrLf & _
          DumpConnState("Cn", Cn) & vbCrLf & _
          DumpConnState("POSConnection", POSCn) & vbCrLf

    ' ADODB provider/SQL errors (very important)
    Msg = Msg & vbCrLf & "---- ADODB Errors ----" & vbCrLf
    Msg = Msg & DumpAdoErrors("Cn", Cn) & vbCrLf
    Msg = Msg & DumpAdoErrors("POSConnection", POSCn) & vbCrLf

    ' ⁄—÷ ··„” Œœ„ + («Œ Ì«—Ì) ﬂ «»… ·„·› ·ÊÃ
    MsgBox Msg, vbCritical, "ConnectionFirst - Detailed Error"
    
    ' ·Ê  Õ»  ”Ã· ›Ì „·›:
    'AppendToLogFile App.Path & "\Logs\app_errors.log", msg
End Sub

Private Function DumpConnState(ByVal Name As String, ByVal c As ADODB.Connection) As String
    On Error Resume Next
    If c Is Nothing Then
        DumpConnState = Name & ": (Nothing)"
        Exit Function
    End If
    
    DumpConnState = Name & ": State=" & c.State
    If c.State = adStateOpen Then
        DumpConnState = DumpConnState & ", Provider=" & c.Provider
    End If
End Function

Private Function DumpAdoErrors(ByVal Name As String, ByVal c As ADODB.Connection) As String
    On Error Resume Next
    
    Dim s As String, i As Long
    If c Is Nothing Then
        DumpAdoErrors = Name & ": (Nothing)"
        Exit Function
    End If
    
    If c.Errors Is Nothing Then
        DumpAdoErrors = Name & ": (No Errors collection)"
        Exit Function
    End If
    
    If c.Errors.Count = 0 Then
        DumpAdoErrors = Name & ": (No provider errors)"
        Exit Function
    End If
    
    s = Name & " Errors.Count=" & c.Errors.Count & vbCrLf
    For i = 0 To c.Errors.Count - 1
        s = s & "  [" & i & "] Number=" & c.Errors(i).Number & _
                ", NativeError=" & c.Errors(i).NativeError & vbCrLf & _
                "      SQLState=" & c.Errors(i).SQLState & vbCrLf & _
                "      Source=" & c.Errors(i).Source & vbCrLf & _
                "      Desc=" & c.Errors(i).Description & vbCrLf
    Next i
    
    DumpAdoErrors = s
End Function

Public Sub AppendToLogFile(ByVal FilePath As String, ByVal Text As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open FilePath For Append As #f
    Print #f, String$(70, "-")
    Print #f, Now & " | " & Text
    Close #f
End Sub

