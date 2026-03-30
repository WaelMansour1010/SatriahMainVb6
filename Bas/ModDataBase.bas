Attribute VB_Name = "ModDataBase"
Option Explicit



Public Enum Key_Type
    KeyPrimary = adKeyPrimary
    KeyUnique = adKeyUnique
End Enum

Public SessionD As Double

    Private m_UserName  As String
    Private m_UserDomain As String
    Private m_ComputerName As String
Private m_macAddress As String
Private m_ObjVmI As Object
Private m_NetworkAdapterVM As Object
Private Property Get VmI() As Object
If m_ObjVmI Is Nothing Then
Set m_ObjVmI = GetObject("winmgmts:\\.\root\cimv2")
End If
   Set VmI = m_ObjVmI
End Property

Private Property Get NetworkAdapterVM() As Object
    If m_NetworkAdapterVM Is Nothing Then
        Set m_NetworkAdapterVM = VmI.ExecQuery("SELECT * FROM " & _
           "Win32_NetworkAdapterConfiguration " & _
           "WHERE IPEnabled = True")
    End If
    Set NetworkAdapterVM = m_NetworkAdapterVM
End Property

Sub CreateRelationship(strDBPath As String, _
                       strForeignTbl As String, _
                       strRelName As String, _
                       strFTKey As String, _
                       strRelatedTbl As String, _
                       strRTKey As String)
    Dim catDB As ADOX.Catalog
    Dim tbl As ADOX.table
    Dim key As ADOX.key

    Set catDB = New ADOX.Catalog
    ' Open the catalog.
    catDB.ActiveConnection = Cn ' "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source =" & strDBPath

    Set key = New ADOX.key

    ' Create the foreign key to define the relationship.
    With key
        ' Specify name for the relationship in the Keys collection.
        .Name = strRelName
        ' Specify the related table's name.
        .RelatedTable = strRelatedTbl
        .type = adKeyForeign
        ' Add the foreign key field to the Columns collection.
        .Columns.Append strFTKey
        ' Specify the field the foreign key is related to.
        .Columns(strFTKey).RelatedColumn = strRTKey
    End With

    Set tbl = New ADOX.table
    ' Open the table and add the foreign key.
    Set tbl = catDB.Tables(strForeignTbl)
    tbl.keys.Append key

    Set catDB = Nothing
End Sub

Public Function db_createOrUpdateFuctionSQL(qry_name As String, _
                                            StrSQL As String)
 
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column
    Dim TableCount As Integer
    'On Error GoTo ErrTrap
 
    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat

    For TableCount = 0 To XCat.Procedures.count - 1

        If qry_name = XCat.Procedures(TableCount).Name Or qry_name & ";0" = XCat.Procedures(TableCount).Name Then
         
            Cn.Execute "drop  FUNCTION " & qry_name
            GoTo ll
            '  Exit Function
        End If

    Next TableCount

ll:
    'StrSQL = "create view " & qry_name & " as  " & StrSQL
    Cn.Execute StrSQL
ErrTrap:

End Function

Public Function db_createOrUpdateviewSQL(ByVal qry_name As String, _
                                         ByVal StrSQL As String) As Boolean
    On Error GoTo ErrTrap

    Dim viewFullName As String
    Dim viewObjName As String
    Dim sqlCreateStub As String
    Dim sqlAlter As String

    '«”„ «·›ÌÊ »«·‹ schema (·· ‰›Ì–)
    viewFullName = "dbo.[" & Replace(qry_name, "]", "]]") & "]"

    '«”„ «·›ÌÊ ﬂ‰’ œ«Œ· OBJECT_ID (»œÊ‰ [])
    viewObjName = "dbo." & qry_name
    viewObjName = Replace(viewObjName, "'", "''")

    '1) ·Ê „‘ „ÊÃÊœ…: «⁄„· Stub View
    sqlCreateStub = ""
    sqlCreateStub = sqlCreateStub & "IF OBJECT_ID(N'" & viewObjName & "', N'V') IS NULL" & vbCrLf
    sqlCreateStub = sqlCreateStub & "BEGIN" & vbCrLf
    sqlCreateStub = sqlCreateStub & "    EXEC(N'CREATE VIEW " & viewFullName & " AS SELECT 1 AS Dummy WHERE 1=0');" & vbCrLf
    sqlCreateStub = sqlCreateStub & "END"
    Cn.Execute sqlCreateStub

    '2) »⁄œ ﬂœÂ ALTER VIEW »«·‹ SQL «·ÕﬁÌﬁÌ
    sqlAlter = "ALTER VIEW " & viewFullName & " AS " & vbCrLf & StrSQL
    Cn.Execute sqlAlter

    db_createOrUpdateviewSQL = True
    Exit Function

ErrTrap:
    db_createOrUpdateviewSQL = False
    '·Ê  Õ»  ŸÂ— ”»» «·Œÿ√:
    'MsgBox "Create/Update View failed: " & Err.Description, vbCritical
End Function

Public Function db_createOrUpdateviewSQLaaa(ByVal qry_name As String, _
                                         ByVal StrSQL As String) As Boolean
    On Error GoTo ErrTrap

    Dim vSql As String
    Dim safeName As String

    '·Ê » ” Œœ„ dbo œ«Ì„«
    safeName = "dbo." & Replace(qry_name, "]", "]]")

    '1) ‰Õ«Ê· CREATE OR ALTER (SQL Server 2016 SP1+ €«·»«)
    vSql = "CREATE OR ALTER VIEW [" & safeName & "] AS " & vbCrLf & StrSQL
    Cn.Execute vSql
    db_createOrUpdateviewSQLaaa = True
    Exit Function

ErrTrap:
    '2) ·Ê ›‘· (€«·»« «·”Ì—›— ﬁœÌ„) ‰⁄„· DROP/CREATE »ÿ—Ìﬁ… ¬„‰…
    On Error GoTo ErrTrap2

    vSql = ""
    vSql = vSql & "IF OBJECT_ID(N'" & Replace(safeName, "'", "''") & "', N'V') IS NOT NULL" & vbCrLf
    vSql = vSql & "    DROP VIEW [" & safeName & "];" & vbCrLf
    vSql = vSql & "EXEC(N'CREATE VIEW [" & Replace(safeName, "]", "]]") & "] AS " & Replace(StrSQL, "'", "''") & "');"

    Cn.Execute vSql
    db_createOrUpdateviewSQLaaa = True
    Exit Function

ErrTrap2:
    db_createOrUpdateviewSQLaaa = False
    '·Ê  Õ»  —„Ì «·Œÿ√ ··Ê«ÃÂ…:
    'MsgBox "Create/Update View failed: " & Err.Description, vbCritical
End Function

Public Function AddSessonData(Optional PointID As Integer, _
                              Optional ShiftID As Integer, _
                              Optional CashierID As Integer, _
                              Optional ShiftFrom As String, _
                              Optional ShiftTo As String, _
                              Optional BoxBalnce As Double, _
                              Optional OHDA As Double, _
                              Optional LoginDateTime As Date, _
                              Optional LogOutDateTime As Date)
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
  '  Exit Function
    SessionD = CStr(new_id("TblSessions", "SessionD", "", True))
    StrSQL = "select * from  TblSessions  where SessionD=-1"
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    rs.AddNew
    rs("SessionD").value = SessionD
    rs("PointID").value = PointID
    rs("ShiftID").value = ShiftID
    rs("CashierID").value = CashierID
    rs("ShiftFrom").value = ShiftFrom
    rs("ShiftTo").value = ShiftTo
     
    rs("BoxBalnce").value = val(BoxBalnce)
    rs("Ohda").value = val(OHDA)

    If Not IsNull(LoginDateTime) Then
        rs("LoginDateTime").value = LoginDateTime
    End If

    If Not IsNull(LogOutDateTime) Then
        rs("LogOutDateTime").value = LogOutDateTime
    End If
     
    rs.update

    rs.Close
     
End Function

Public Function AddToLogFile(Optional UserID As Integer, _
                             Optional NotesType As Integer = -1, _
                             Optional LogDate As Date, _
                             Optional LogTime As Date, _
                             Optional Description As String, _
                             Optional DescriptionE As String, _
                             Optional Remarks As String, _
                             Optional transactiontype As String, _
                             Optional X As String, _
                             Optional Y As String, _
                             Optional NoteSerial As Variant = 0, _
                             Optional NoteSerial1 As String = "")

    On Error GoTo ErrTrap
    Dim StrSQL         As String
    Dim Columns        As String
    Dim Values         As String
    Dim rs             As ADODB.Recordset
    
    Dim sComputerName  As String
    Dim UserName       As String
    Dim UserDomain     As String
    Dim MACAddress     As String
    Dim ConnectionData As String
    'Dim myWMI As Variant
    'Dim myObj As Variant
    Dim Itm            As Variant
    'Set myWMI = GetObject("winmgmts:\\.\root\cimv2")
    'Set myObj = VmI.ExecQuery("SELECT * FROM " & _
    '                 "Win32_NetworkAdapterConfiguration " & _
    '                 "WHERE IPEnabled = True")
    If m_UserName & "" = "" Then
        m_UserName = Environ("USERNAME")
       
    End If
    If m_UserDomain & "" = "" Then
        m_UserDomain = Environ("USERDOMAIN")
    End If
    If m_ComputerName & "" = "" Then
        m_ComputerName = Environ("computername")
    End If
   
    If m_macAddress & "" = "" Then
        For Each Itm In NetworkAdapterVM
            m_macAddress = (Itm.MACAddress)
            Exit For
        Next
    End If
    UserName = m_UserName
    UserDomain = m_UserDomain
    sComputerName = m_ComputerName
    MACAddress = m_macAddress

    If Remarks = "FrmLogIn" Then
        Dim LocalStr As String
        LocalStr = "insert into  TblGroupItemProductLineUsersset (MACAddress ,computername,UserName,UserDomain,ProgramUsername)"
        LocalStr = LocalStr & " Values ( '" & MACAddress & "' ,'" & sComputerName & "','" & UserName & "', '" & UserDomain & "' ,'" & user_name & "')"

        Cn.Execute LocalStr
    End If

    ConnectionData = " UserName : " + UserName + CHR(13) + " user Domain   " + UserDomain + CHR(13) + " Computer Name: " + sComputerName + CHR(13) + " MAC Address  :  " + MACAddress

    NoteSerial = val(NoteSerial)
    If NotesType = 0 Then NotesType = -1
    GetServerdate LogDate, LogTime
    
    ConnectionData = ConnectionData + CHR(13) + " PC Date/Time:" & Now
    ConnectionData = ConnectionData + CHR(13) + " Server Date/Time:" & LogDate & " : " & LogTime

    Columns = "UserID,NotesType,LogDate,LogTime,Description,Descriptione,Remarks,TransactionType,NotesSerial,NotesSerial1,ConnectionData,ProgramUsername,Computername,ComputerUsername"
    
    '  Values = UserID & "," & notesType & "," & SQLDate(LogDate, True) & ",'" & Format(LogTime, "hh:mm:ss") & "','" & description & "','" & DescriptionE & "','" & Remarks & "','" & transactiontype & "','" & NoteSerial & "','" & NoteSerial1 & "','" & ConnectionData & "'"
    
    Values = UserID & "," & NotesType & ",  Getdate()    , Getdate()       , '" & Description & "','" & DescriptionE & "','" & Remarks & "','" & transactiontype & "','" & NoteSerial & "','" & NoteSerial1 & "','" & ConnectionData & "','" & user_name & "','" & sComputerName & "','" & UserName & "'"
 
    StrSQL = "INSERT INTO  LogFile (" & Columns & ")" & " VALUES ( " & Values & " )"
    
    Cn.Execute StrSQL

ErrTrap:
End Function

'add record to table
Public Function add_record_to_table(ByVal tablename As String, _
                                    Columns As String, _
                                    Values As String, _
                                    Optional Search_field As String, _
                                    Optional search_value As Variant) As Boolean
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
    StrSQL = "select * from " & tablename & " where " & Search_field & " = " & search_value
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount = 0 Then
        StrSQL = "INSERT INTO " & tablename & "(" & Columns & ")" & " VALUES ( " & Values & " )"
        Cn.Execute StrSQL
        add_record_to_table = True
    Else
        add_record_to_table = False
    End If

ErrTrap:
End Function

'Update record to table
Public Function update_record_to_table(ByVal tablename As String, _
                                       Columns As String, _
                                       Values As String, _
                                       Optional Search_field As String, _
                                       Optional search_value As Integer) As Boolean
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

    If IsNumeric(Values) Then
        StrSQL = "update  " & tablename & " set " & Columns & "=" & Values
    Else
        StrSQL = "update  " & tablename & " set " & Columns & "='" & Values & "'"
    End If

    StrSQL = StrSQL + " where " & Search_field & " = " & search_value
    ' rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    ' If rs.RecordCount = 0 Then
    'StrSQL = "INSERT INTO " & tablename & "(" & Columns & ")" & _
    '" VALUES ( " & Values & " )"
    Cn.Execute StrSQL
 
ErrTrap:
End Function

Public Function DB_CreateTable(ByVal tablename As String, _
                               Optional PK As Boolean = False, _
                               Optional PK_NAME As String, _
                               Optional identity As Boolean = False, _
                               Optional PKConst As String = "") As Boolean
    '    Dim StrMSG As String
    '    Dim XCat As ADOX.Catalog
    '    Dim XTable As ADOX.table
    '    Dim XCol As ADOX.Column
    '    Dim TableCount As Integer
    '    On Error GoTo ErrTrap
    '    DB_CreateTable = False
    '    Set XCat = New ADOX.Catalog
    '    Set XTable = New ADOX.table
    '    Set XCat.ActiveConnection = Cn
    '    Set XTable.ParentCatalog = XCat
    '-------------------------------Create New Table---------------------------------------------

    '    For TableCount = 0 To XCat.Tables.count - 1
    '
    '        If tablename = XCat.Tables(TableCount).Name Then
    '            DB_CreateTable = False
    '            Exit Function
    '        End If
    '
    '    Next TableCount

    On Error GoTo hErr
    Dim sql As String
    sql = ""
    sql = sql & "IF NOT EXISTS "
    sql = sql & "( "
    sql = sql & "    SELECT * "
    sql = sql & "    FROM sysobjects "
    sql = sql & "    WHERE name = '" & tablename & "' "
    sql = sql & "          AND xtype = 'U' "
    sql = sql & ") "
    sql = sql & " BEGIN "
    sql = sql & "    CREATE TABLE " & tablename
    sql = sql & "    ( "
    If InStr(1, PK_NAME, ",", vbTextCompare) > 0 Then
        sql = sql & PK_NAME
        If PKConst <> "" Then
            sql = sql & " , "
            sql = sql & " Constraint [PK_INDX_" & tablename & "]"
            sql = sql & "   PRIMARY KEY CLUSTERED (" & PKConst
            sql = sql & "                    ) ON [PRIMARY]"
        End If
    Else
        If PK = True And identity = True Then
            sql = sql & "     " & PK_NAME & " int IDENTITY(1,1) NOT NULL PRIMARY KEY "
        ElseIf PK = True And identity = False Then
            sql = sql & "     " & PK_NAME & " int  NOT NULL PRIMARY KEY "
        ElseIf PK = False And identity = True Then
            sql = sql & "     " & PK_NAME & " int IDENTITY(1,1) NOT NULL   "
        ElseIf PK = False And identity = False Then
            sql = sql & "     " & PK_NAME & " int  NOT NULL   "
        End If
    End If
  
    sql = sql & "    ) ON [PRIMARY]; END"

    Cn.Execute sql
    DB_CreateTable = True
    Exit Function
hErr:
    DB_CreateTable = False

End Function

'------------------------- create Table in database ---------------

'
'Public Function DB_CreateTable(ByVal tablename As String, _
'                               Optional PK As Boolean = False, _
'                               Optional PK_NAME As String, _
'                               Optional identity As Boolean = False) As Boolean
'    Dim StrMSG As String
'    Dim XCat As ADOX.Catalog
'    Dim XTable As ADOX.table
'    Dim XCol As ADOX.Column
'    Dim TableCount As Integer
'    On Error GoTo ErrTrap
'    DB_CreateTable = False
'    Set XCat = New ADOX.Catalog
'    Set XTable = New ADOX.table
'    Set XCat.ActiveConnection = Cn
'    Set XTable.ParentCatalog = XCat
'    '-------------------------------Create New Table---------------------------------------------
'
'    For TableCount = 0 To XCat.Tables.count - 1
'
'        If tablename = XCat.Tables(TableCount).Name Then
'            DB_CreateTable = False
'            Exit Function
'        End If
'
'    Next TableCount
'
'    Dim StrSQL As String
'
'    If PK = True And identity = True Then
'        StrSQL = "CREATE TABLE " & tablename & " ( " & PK_NAME & " int IDENTITY(1,1) NOT NULL PRIMARY KEY)"
'    ElseIf PK = True And identity = False Then
'        StrSQL = "CREATE TABLE " & tablename & " ( " & PK_NAME & " int  NOT NULL PRIMARY KEY)"
'    ElseIf PK = False And identity = True Then
'        StrSQL = "CREATE TABLE " & tablename & " ( " & PK_NAME & " int IDENTITY(1,1) NOT NULL  )"
'    ElseIf PK = False And identity = False Then
'        StrSQL = "CREATE TABLE " & tablename & " ( " & PK_NAME & " int  NOT NULL  )"
'    End If
'
'    Cn.Execute StrSQL
'
'    DB_CreateTable = True
'    Exit Function
'
'    Dim KeyName As String
'
'    XTable.Name = tablename
'
'    KeyName = "Tempcol"
'
'    If PK = True Then
'        Dim XKey As ADOX.Key
'        Set XKey = New ADOX.Key
'        KeyName = PK_NAME
'        XKey.Name = KeyName
'        XKey.type = adKeyPrimary
'        XKey.Columns.Append KeyName
'
'        XTable.keys.Append XKey
'    End If
'
'    XTable.Columns.Append KeyName, adInteger
'
'    XCat.Tables.Append XTable
'    DB_CreateTable = True
'    Set XCol = Nothing
'    Set XTable = Nothing
'    Set XCat = Nothing
'    Exit Function
'ErrTrap:
'
'    Select Case Err.Number
'
'        Case "-2147217857"
'            'StrMSG = "⁄›Ê« ·ﬁœ ”»ﬁ ≈‰‘«¡" & " Â–« «·ÃÊœ· „‰ ﬁ»·"
'            'MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
'            DB_CreateTable = False
'    End Select
'
'End Function

Public Function DB_updateField(ByVal tablename As String, _
                               FieldName As String, _
                               FieldType As String)
   
    Dim idxNew As New ADOX.Index
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column
    Dim FeildsNum As Integer
    'On Error GoTo ErrTrap
 
    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat
    '-------------------------------update field data type---------------------------------------------

    Set XTable = XCat(tablename)
    Set XCol = New ADOX.Column

    For FeildsNum = 0 To XTable.Columns.count - 1

        If XTable.Columns(FeildsNum).Name = FieldName Then
         
           GoTo ll
        End If

    Next FeildsNum
 Exit Function
ll:
    Dim StrSQL As String
    StrSQL = "ALTER TABLE " & tablename & " " & " ALTER column  " & FieldName & "  " & FieldType & " ;"
    Cn.Execute StrSQL

End Function

Public Function DB_CreateField(ByVal tablename As String, _
                               FieldName As String, _
                               Optional FieldType As ADODB.DataTypeEnum, _
                               Optional FiledAttrib As ColumnAttributesEnum, _
                               Optional FieldSize As Integer, _
                               Optional DefaultValue As String, _
                               Optional FieldDescription As String, _
                               Optional Required As Boolean, _
                               Optional ZeroLength As Boolean, _
                               Optional KeyType As Key_Type, _
                               Optional AUTONO As Boolean = False) As Boolean
   
    Dim idxNew As New ADOX.Index
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column
    Dim FeildsNum As Integer
    'On Error GoTo ErrTrap
    DB_CreateField = False
    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat
    '-------------------------------Create New field---------------------------------------------

    Set XTable = XCat(tablename)
    Set XCol = New ADOX.Column

    For FeildsNum = 0 To XTable.Columns.count - 1

        If XTable.Columns(FeildsNum).Name = FieldName Then
            DB_CreateField = False
            Exit Function
        End If

    Next FeildsNum

    XCol.Name = FieldName

    If FieldType = 0 Then
        FieldType = adVarWChar
    End If

    XCol.type = FieldType

    If FieldType <> adBoolean Then
        'XCol.Attributes = adColNullable ' FiledAttrib
    End If

    XCol.Attributes = FiledAttrib

    If FieldType = adVarChar Or FieldType = adVarWChar Or FieldType = adWChar Then
        If FieldSize = 0 Then
            FieldSize = 255
        End If

        XCol.DefinedSize = FieldSize
    End If

    If FieldType = adDecimal Then
        XCol.Precision = 18
        XCol.NumericScale = 2
    End If

    If AUTONO = True Then
        'XCol.Properties ("Autoincrement")
    End If
    On Error Resume Next
    XTable.Columns.Append XCol

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        XCol.Properties("Description").value = FieldDescription
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        XCol.Properties("Default").value = DefaultValue
        XTable.Columns(FieldName).Properties("Jet OLEDB:Allow Zero Length").value = ZeroLength
    End If

    'XTable.Columns(FieldName).Properties("Requied").Value = Required
    If KeyType <> 0 Then

        Select Case KeyType

            Case adKeyPrimary
                '    XTable.Columns(FieldName).Properties("Requied").value = Required
                ' XTable.Keys.Append "PrimaryKey", KeyType, FieldName
                idxNew.Name = "NumIndex"
                idxNew.Columns.Append FieldName
                idxNew.PrimaryKey = True
                idxNew.Unique = True
                XTable.Indexes.Append idxNew
    
            Case adKeyForeign
                XTable.keys.Append "ForeignKey", KeyType, FieldName

            Case adKeyUnique
                XTable.keys.Append "Unique", KeyType, FieldName
        End Select

    End If

    DB_CreateField = True
    Set XCol = Nothing
    Set XTable = Nothing
    Set XCat = Nothing
    Exit Function
ErrTrap:

    Select Case Err.Number

        Case "-2147217858"
            StrMSG = "⁄›Ê« ·ﬁœ ”»ﬁ ≈‰‘«¡" & " Â–« «·Õﬁ· „‰ ﬁ»·"
            'MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
            DB_CreateField = False

        Case "-2147217903"
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ ≈‰‘«¡" & " „› «Õ ›—⁄Ì ›Ì Â–… «·œ«·…"
            'MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
            DB_CreateField = False

        Case "-2147217767"
            StrMSG = "⁄›Ê« „› «Õ" & " PrimaryKey" & "„ÊÃÊœ »«·›⁄·"
            'MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
            DB_CreateField = False
    End Select

End Function

Public Function DB_PrimaryKey(ByVal tablename As String, _
                              KeyName As String) As Boolean
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column
    Dim XKey As ADOX.key

    On Error GoTo ErrTrap
    DB_PrimaryKey = False
    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat
    '-------------------------------Create Primary Key --------------------------------------------
    Set XTable = XCat(tablename)
    Set XCol = New ADOX.Column

    Set XKey = New ADOX.key

    XKey.Name = KeyName
    XKey.type = adKeyPrimary
    XKey.Columns.Append KeyName
    XTable.keys.Append XKey
    XTable.Columns(KeyName).Properties(0).value = True
    Set XKey = Nothing
    Set XCol = Nothing
    Set XTable = Nothing
    Set XCat = Nothing

    Exit Function
ErrTrap:

    Select Case Err.Number

        Case "-2147217767"
            StrMSG = "⁄›Ê« „› «Õ" & " PrimaryKey" & "„ÊÃÊœ »«·›⁄·"
            ' MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
            DB_PrimaryKey = False
    End Select

End Function

Public Function DB_Relations(ByVal PrimaryTable As String, _
                             PrimaryKey As String, _
                             ForeignTable As String, _
                             ForeignKey As String, _
                             Optional IntDeleteRule As RuleEnum = RuleEnum.adRINone, _
                             Optional IntUpdateRule As RuleEnum = RuleEnum.adRINone)
 
    'DB_Relations "TransactionTypes", "Transaction_Type", "Transactions", "Transaction_Type"
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column
    Dim FKey As ADOX.key
    On Error GoTo ErrTrap
    'DB_PrimaryKey = False
    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat
    Set XTable = XCat(ForeignTable)
    '-------------------------------Create Primary Key --------------------------------------------
    'XTable.Name = ForeignTable
    Set XCol = New ADOX.Column
    Set FKey = New ADOX.key
    Dim ROWN As Integer

    For ROWN = 0 To XTable.keys.count - 1

        If XTable.keys(ROWN).Name = ForeignKey Then
            Exit Function
        End If

    Next ROWN

    FKey.Name = ForeignKey
    FKey.type = adKeyForeign

    FKey.RelatedTable = PrimaryTable
    FKey.Columns.Append ForeignKey
    FKey.Columns(ForeignKey).RelatedColumn = PrimaryKey

    FKey.UpdateRule = IntUpdateRule
    FKey.DeleteRule = IntDeleteRule

    XTable.keys.Append FKey
    '------------------------------------------------------------------------------------------------------
    Set FKey = Nothing
    Set XCol = Nothing
    Set XTable = Nothing
    Set XCat = Nothing
    Exit Function
ErrTrap:
    StrMSG = "··œ⁄„ «·›‰Ì √ ’· »«·‘—ﬂ… «·„‰ Ã…"
    'MsgBox StrMSG, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation, App.Title
End Function

Public Sub rln()
    Dim tbl As ADOX.table
    Dim Cat As ADOX.Catalog
    Dim fk As ADOX.key
    Set tbl = New ADOX.table
    Set Cat = New ADOX.Catalog
    'Set the ParentCatalog property on the Table
    ' to expose the database-specific properties.
    Set Cat.ActiveConnection = Cn

    Set tbl.ParentCatalog = Cat
    Set tbl = Cat("Letters")

    tbl.Name = "Letters"
    'Add the Columns to the table.
    tbl.Columns.Append "OrderID", adInteger
    tbl.Columns("OrderID").Properties("AutoIncrement") = True
    tbl.Columns.Append "Num", adWChar, 5

    'Create the primary key.
    tbl.keys.Append "PK_Orders", adKeyPrimary, "OrderID"
    'Create the foreign key.
    'You must explicitly create the key this way to set the
    ' DeleteRule property.
    Set fk = New ADOX.key
    fk.Name = "Num"
    fk.type = adKeyForeign
    fk.RelatedTable = "Customers"
    fk.Columns.Append "Num"
    fk.Columns("Num").RelatedColumn = "CustomerID"
    fk.UpdateRule = adRICascade
    tbl.keys.Append fk
    'Create the indexes.
    tbl.Indexes.Append "IDX_Orders_Customers", "CustomerID"
    'Add the table to the database.
    Cat.Tables.Append tbl

End Sub

Public Function check_relation_exist(parent_table As String, _
                                     parent_table_pk As String, _
                                     child_table As String, _
                                     child_table_fk As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = " SELECT " & " K_Table = FK.TABLE_NAME," & " FK_Column = CU.COLUMN_NAME," & " PK_Table = PK.TABLE_NAME," & " PK_Column = PT.COLUMN_NAME," & " CONSTRAINT_NAME = C.CONSTRAINT_NAME " & " FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS C " & " INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS FK ON C.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & " INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS PK ON C.UNIQUE_CONSTRAINT_NAME = PK.CONSTRAINT_NAME " & " INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE CU ON C.CONSTRAINT_NAME = CU.CONSTRAINT_NAME " & " INNER JOIN ( " & " SELECT i1.TABLE_NAME, i2.COLUMN_NAME " & " FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS i1 " & " INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE i2 ON i1.CONSTRAINT_NAME = i2.CONSTRAINT_NAME " & ") PT ON PT.TABLE_NAME = PK.TABLE_NAME " & "where     CU.COLUMN_NAME='" & child_table_fk & "'  and PT.COLUMN_NAME='" & parent_table_pk & "' and FK.TABLE_NAME='" & child_table & "' and PK.TABLE_NAME='" & parent_table & "' "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        check_relation_exist = True
    Else
        check_relation_exist = False
    End If

End Function

Public Function db_createRelationSQL(parent_table As String, _
                                     parent_table_pk As String, _
                                     child_table As String, _
                                     child_table_fk As String)

    If check_relation_exist(parent_table, parent_table_pk, child_table, child_table_fk) = True Then
        Exit Function
    End If

    Dim relation_name As String
    relation_name = "FK_" & child_table & "_" & parent_table
    Dim StrSQL As String
    StrSQL = "ALTER TABLE " & child_table & "  " & " ADD CONSTRAINT " & relation_name & " " & "    FOREIGN KEY  (" & child_table_fk & ") REFERENCES  " & parent_table & "(" & parent_table_pk & ") ON DELETE CASCADE ON update CASCADE"
    Cn.Execute StrSQL
ErrTrap:
End Function

Public Function db_deleteRelationSQL(parent_table As String, _
                                     parent_table_pk As String, _
                                     child_table As String, _
                                     child_table_fk As String)

    If check_relation_exist(parent_table, parent_table_pk, child_table, child_table_fk) = False Then
        Exit Function
    End If
  
    Dim relation_name As String
    relation_name = "FK_" & child_table & "_" & parent_table
    Dim StrSQL As String

    StrSQL = " ALTER TABLE  " & child_table & " drop CONSTRAINT " & relation_name
    Cn.Execute StrSQL
    Exit Function

    Dim New_View As String
    New_View = "SELECT  fk.constraint_name,CU.COLUMN_NAME as FKfiled ,PT.COLUMN_NAME as PKfield  ,FK.TABLE_NAME as FKtable,PK.TABLE_NAME as pktable  FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS C " & " INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS FK ON C.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & "  INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS PK ON C.UNIQUE_CONSTRAINT_NAME = PK.CONSTRAINT_NAME " & " INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE CU ON C.CONSTRAINT_NAME = CU.CONSTRAINT_NAME " & " INNER JOIN (  SELECT i1.TABLE_NAME, i2.COLUMN_NAME  FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS i1 " & " INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE i2 ON i1.CONSTRAINT_NAME = i2.CONSTRAINT_NAME ) PT ON PT.TABLE_NAME = PK.TABLE_NAME " & " where     CU.COLUMN_NAME='" & child_table_fk & "'  and PT.COLUMN_NAME='" & parent_table_pk & "' and FK.TABLE_NAME='" & child_table & "' and PK.TABLE_NAME='" & parent_table & "'"
    db_createOrUpdateviewSQL "relation_all_info", New_View

    'Dim StrSQL As String
    StrSQL = "delete relation_all_info where FKfiled='" & child_table_fk & "' and  PKfield='" & parent_table_pk & "' and fktable='" & child_table & "' and pktable='" & parent_table_pk & "'"

    Cn.Execute StrSQL
ErrTrap:
End Function
 
Public Function DB_CreateView(ByVal QRYNAME As String, _
                              ByVal StrSQL As String)
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    'Dim XTable As ADOX.Table
    'Dim XCol As ADOX.Column

    Dim XView As ADOX.View
    Dim SqlCmd As ADODB.Command
    Dim ViewNum As Integer
    On Error GoTo ErrTrap
    'DB_CreateTable = False
    Set XCat = New ADOX.Catalog
    'Set XTable = New ADOX.Table
    Set XCat.ActiveConnection = Cn

    'Set XTable.ParentCatalog = XCat
    '-------------------------------Create New Query---------------------------------------------
    For ViewNum = 0 To XCat.Views.count - 1

        If XCat.Views(ViewNum).Name = QRYNAME Then
            Exit Function
        End If

    Next ViewNum

    Set SqlCmd = New ADODB.Command  'XCat.Views("qry").Command
    SqlCmd.CommandText = StrSQL
    'Set XCat.Views("qry").Command = SQLCmd
    On Error GoTo ErrTrap

    XCat.Views.Append QRYNAME, SqlCmd
    Exit Function
ErrTrap:

    'XTable.Name = TableName
    'XCat.Tables.Append XTable
    'XView.Command = cmd

End Function

Public Function DB_UpDateView(ByVal QryExist As String, _
                              ByVal QryNew As String)
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim rs As ADODB.Recordset
    Dim RsScd As ADODB.Recordset
    Dim i As Integer
    Dim ii As Integer
    Dim XView As ADOX.View
    Dim SqlCmd As ADODB.Command
    Dim BolUpdate As Boolean
    On Error GoTo ErrTrap

    Set XCat = New ADOX.Catalog
    Set XCat.ActiveConnection = Cn
    Set rs = New ADODB.Recordset
    Set RsScd = New ADODB.Recordset
    RsScd.Open QryNew, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    rs.Open QryExist, Cn, adOpenForwardOnly, adLockReadOnly

    For ii = 0 To RsScd.Fields.count - 1
        For i = 0 To rs.Fields.count - 1

            If rs.Fields(i).Name = RsScd.Fields(ii).Name Then
                GoTo Nxt
                'Else
            End If

        Next i

        BolUpdate = True
        Exit For
Nxt:

    Next ii

    If BolUpdate = True Then
        Set SqlCmd = XCat.Views(QryExist).Command
        SqlCmd.CommandText = QryNew
        Set XCat.Views(QryExist).Command = SqlCmd
        Set SqlCmd = Nothing
    End If

    rs.Close
    RsScd.Close
    Set RsScd = Nothing
    Set rs = Nothing
    Set XView = Nothing
    Set XCat = Nothing
    Exit Function
ErrTrap:
End Function

Public Sub FieldDescr(ByVal tablename As String, _
                      FieldName As String)
    Dim StrMSG As String
    Dim XCat As ADOX.Catalog
    Dim XTable As ADOX.table
    Dim XCol As ADOX.Column

    On Error GoTo ErrTrap

    Set XCat = New ADOX.Catalog
    Set XTable = New ADOX.table
    Set XCat.ActiveConnection = Cn
    Set XTable.ParentCatalog = XCat
    '-------------------------------Create New field---------------------------------------------

    Set XTable = XCat(tablename)
    Set XCol = New ADOX.Column
    Dim i As Integer

    XCol.Name = FieldName
    Debug.Print "---------------------" & FieldName & "-----------------------"

    For i = 0 To 10
        Debug.Print XTable.Columns(FieldName).Properties(i).Attributes & "***" & XTable.Columns(FieldName).Properties(i).Name & "***" & XTable.Columns(FieldName).Properties(i).type & "***" & XTable.Columns(FieldName).Properties(i).value
        Debug.Print "---New Row -----------------------------------------"
    Next

    Set XCol = Nothing
    Set XTable = Nothing
    Set XCat = Nothing

ErrTrap:
End Sub

Public Function DB_DelView(ByVal tablename As String) As Boolean
    On Error GoTo ErrTrap
    Dim XCat As ADOX.Catalog
    Dim ViewNum As Integer
    Set XCat = New ADOX.Catalog
    Set XCat.ActiveConnection = Cn

    For ViewNum = 0 To XCat.Views.count - 1

        If XCat.Views(ViewNum).Name = tablename Then
            XCat.Views.delete tablename
            DB_DelView = True
            Exit Function
        End If

    Next ViewNum

    Set XCat = Nothing
    Exit Function
ErrTrap:
    DB_DelView = False
End Function

Public Function DB_UpdateSqlQry(QRYNAME As String, _
                                StrSQL As String)
    Dim XCat As New ADOX.Catalog
    Dim SqlCmd As ADODB.Command
    Dim StrTempOrg As String
    Dim StrTempUpdate As String
    Dim ExitSql As String

    XCat.ActiveConnection = Cn
    Set SqlCmd = XCat.Views(QRYNAME).Command
    StrTempOrg = Replace(SqlCmd.CommandText, CHR(10), vbNullString, , , vbBinaryCompare)
    StrTempOrg = Replace(StrTempOrg, CHR(13), vbNullString, , , vbBinaryCompare)
    StrTempOrg = Replace(StrTempOrg, CHR(32), vbNullString, , , vbBinaryCompare)

    StrTempUpdate = Replace(StrSQL, CHR(10), vbNullString, , , vbBinaryCompare)
    StrTempUpdate = Replace(StrTempUpdate, CHR(13), vbNullString, , , vbBinaryCompare)
    StrTempUpdate = Replace(StrTempUpdate, CHR(32), vbNullString, , , vbBinaryCompare)
 
    If StrComp(StrTempOrg, StrTempUpdate, vbTextCompare) <> 0 Then
        ExitSql = SqlCmd.CommandText
        SqlCmd.CommandText = StrSQL
        Set XCat.Views(QRYNAME).Command = SqlCmd
    End If

    Set SqlCmd = Nothing
End Function
