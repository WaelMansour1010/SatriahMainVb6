Attribute VB_Name = "MyFunctions"

Public Sub UpdateFiles(ByVal POSlServer As String, ByVal POSDb As String, ByVal mTableName As String, Optional ByVal mFieldName As String = "Id", Optional ByVal mWhere As String = "")
    
   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   Dim ss As String
    
    
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
 
            
             
            s = ""
             
            Dim mPosD As String
            Dim mServerD As String
             mPosD = "[" & POSlServer & "]" & "." & POSDb & ".dbo."
             mServerD = "[" & SysSQLServerName & "]" & "." & ServerDb & ".dbo."
             mServerD = ServerDb & ".dbo."
            
         
           ' Text4 = s
           ' Exit Sub
            s = GetSql(mServerD, mPosD, mTableName, mFieldName)
            Cn.Execute s
           ' MsgBox "Step 4"
            
            
           '  MsgBox " „ ‰ﬁ· »Ì«‰«  «·»Ì«‰« "
           '  cmdUdateFiles.Enabled = False

End Sub



Public Sub UpdateFilesFromPos(ByVal POSlServer As String, ByVal POSDb As String, ByVal mTableName As String, Optional ByVal mFieldName As String = "Id", Optional ByVal mWhere As String = "", Optional ByVal mPOSlServer As String = "")
    
   Dim NoOFItem_POS As Double
   Dim NoOFItem_Server As Double
   
   Dim Rs3 As New ADODB.Recordset
   Dim MaxItem_POS As Double
   Dim MaxItem_Server As Double
   Dim ss As String
    
    
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
             
            Dim mPosD As String
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



Public Function GetSql2(ByVal POSlServer As String, ByVal POSDb As String, ByVal mTableName As String, Optional ByVal mFieldName As String = "Id", Optional ByVal mWhere As String = "") As String
    
        Dim s As String
    s = "Select * from " & mTableName & " Where 1 = -1  " ' & mWhere
    Dim rs As New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim mName As String
    Dim mName2 As String
    Dim ii As Long
    mName = ""
    mName2 = ""
    ii = 0
        For ii = 0 To rs.Fields.Count - 1
       If mName = "" Then
            If mTableName = "Transaction_Details" Or mTableName = "TransactionValueAdded" Then
                         If UCase(rs(ii).name) = "ID" Then
                             GoTo NextCol
                         End If
             End If
        mName = mName & rs(ii).name
        mName2 = mName2 & "TT." & Trim(rs(ii).name)
        Else
            If rs(ii).name <> "MainOperationID" Then
                If mTableName = "Transaction_Details" Or mTableName = "TransactionValueAdded" Then
                    If UCase(rs(ii).name) = "ID" Then
                        GoTo NextCol
                    End If
                End If
                If ii = 7 Or ii = 14 Or ii = 21 Or ii = 28 Or ii = 35 Or ii = 42 Or ii = 49 Or ii = 56 Or ii = 66 Or ii = 77 Or ii = 82 Or ii = 90 Or ii = 95 Or ii = 100 Or ii = 107 Or ii = 113 Or ii = 117 Or ii = 125 Or ii = 130 Or ii = 140 Or ii = 149 Or ii = 156 Or ii = 162 Or ii = 170 Or ii = 177 Or ii = 182 Or ii = 190 Or ii = 195 Or ii = 202 Or ii = 210 Or ii = 215 Or ii = 221 Or ii = 230 Or ii = 238 Or ii = 244 Or ii = 250 Or ii = 257 Or ii = 262 Or ii = 270 Or ii = 277 Or ii = 282 Or ii = 290 Or ii = 297 Or ii = 305 Or ii = 312 Or ii = 320 Or ii = 328 Or ii = 335 Or ii = 342 Then
                    mName = mName & "," & rs(ii).name & vbNewLine
                    mName2 = mName2 & ",TT." & rs(ii).name & vbNewLine
                Else
                    mName = mName & "," & rs(ii).name
                    mName2 = mName2 & ",TT." & rs(ii).name

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

Public Function GetSql(ByVal POSlServer As String, ByVal POSDb As String, ByVal mTableName As String, Optional ByVal mFieldName As String = "Id", Optional ByVal mWhere As String = "") As String
    Dim s As String
    s = "Select * from " & mTableName & " Where 1 = -1  " ' & mWhere
    Dim rs As New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim mName As String
    Dim mName2 As String
    Dim ii As Long
    mName = ""
    mName2 = ""
    For ii = 0 To rs.Fields.Count - 1
       If mName = "" Then
        mName = mName & rs(ii).name
        mName2 = mName2 & "T2." & Trim(rs(ii).name)
        Else
            If rs(ii).name <> "MainOperationID" Then
                mName = mName & "," & rs(ii).name & vbNewLine
                mName2 = mName2 & ",T2." & rs(ii).name & vbNewLine
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
    
    GetSql = s
End Function
