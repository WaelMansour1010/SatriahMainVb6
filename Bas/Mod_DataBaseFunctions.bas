Attribute VB_Name = "Mod_DataBaseFunctions"
 
Option Explicit
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Type LastItemTransInfo
    Transactionid As Long
    TransactionSerial As String
    TransactionDate As String
    StrCustomerName As String
    SngItemPrice As Single
    SngItemQty As Single
End Type
Dim Askinterval As String
Dim Askcount    As Integer

Public Sub TerminateRecordset(ByRef rs As ADODB.Recordset)
    On Local Error Resume Next

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            If Not (rs.BOF Or rs.EOF) Then
                If rs.EditMode <> adEditNone Then
                    rs.CancelUpdate
                End If
            End If

            rs.Close
        End If

        Set rs = Nothing
    End If

End Sub

Public Function new_id(tablename As String, _
                       FieldName As String, _
                       str_code As String, _
                       Optional serial As Boolean = False, _
                       Optional StrWhere As String = "") As String
    'This Function to
    'Get the New ID and Serials
    Dim My_SQL  As String
    Dim Lngid   As Long
    Dim Rs_Temp As New ADODB.Recordset

    If serial = False Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            My_SQL = " SELECT Max(Val(Mid$(" & FieldName & " , " & Len(str_code) + 1 & " ,225))) AS max_n "
            My_SQL = My_SQL + " FROM " & tablename & "  "
            My_SQL = My_SQL + " WHERE " & FieldName & "  like  ('" & str_code & "%') "
            Rs_Temp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly ' , adCmdText

            If IsNull(Rs_Temp("max_n").value) Then
                new_id = str_code & "1"
            Else
                new_id = str_code & CStr(val(Rs_Temp("max_n").value) + 1)
            End If

            Set Rs_Temp = Nothing
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            My_SQL = "select max(cast(isnull(" & FieldName & ",0) as int )) as max_n "
            My_SQL = My_SQL + "From " & tablename & ""
            My_SQL = My_SQL + " WHERE " & FieldName & "  like  ('" & str_code & "%') "
            Rs_Temp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

            If IsNull(Rs_Temp("max_n").value) Then
                new_id = str_code & "1"
            Else
                new_id = str_code & CStr(val(Rs_Temp("max_n").value) + 1)
            End If
        End If

    Else

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            My_SQL = " SELECT Max (Val(iif(" & FieldName & " is NULL,0," & FieldName & "))) AS max_n "
            My_SQL = My_SQL + " FROM " & tablename & ""

            If StrWhere <> "" Then
                My_SQL = My_SQL + " Where " & StrWhere
            End If

            Rs_Temp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

            If IsNull(Rs_Temp("max_n").value) Then
                new_id = "1"
            Else
                new_id = CStr(val(Rs_Temp("max_n").value) + 1)
            End If

            Set Rs_Temp = Nothing
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            My_SQL = "select max(cast(isnull(" & FieldName & ",0) as float )) as max_n "
            My_SQL = My_SQL + "From " & tablename & ""

            If StrWhere <> "" Then
                My_SQL = My_SQL + " Where " & StrWhere & " AND isnumeric(" & FieldName & ")=1"
            End If

            Rs_Temp.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly  ' , adCmdText

            If IsNull(Rs_Temp("max_n").value) Then
                new_id = "1"
            Else
                new_id = CStr(val(Rs_Temp("max_n").value) + 1)
            End If

            Set Rs_Temp = Nothing
        End If
    End If

End Function

Public Function Insert(ByVal rs As ADODB.Recordset, _
                       Optional Frm As Form, _
                       Optional Txt, _
                       Optional update As String = "", _
                       Optional types) As Boolean
    Dim i As Integer
    On Error GoTo err_trap

    If update = "" Then rs.AddNew
    If IsMissing(types) Then

        For i = 0 To Frm.Txt.count - 1

            If InStr(1, Frm.Txt(i).Tag, "n") Then rs(i) = IIf(Txt(i) = "", 0, Txt(i))
            If InStr(1, Frm.Txt(i).Tag, "r") Then rs(i) = IIf(Txt(i) = "", 0, Txt(i))
            If InStr(1, Frm.Txt(i).Tag, "d") Then
                If Txt(i) <> "" Then rs(i) = Format(Txt(i), "yyyy/M/d")
            End If

            If InStr(1, Frm.Txt(i).Tag, "s") Then rs(i) = IIf(IsNull(Txt(i)) Or Txt(i) = "", Null, Txt(i))
            Debug.Print rs(i).Name, rs(i).value
        Next i

    Else

        For i = 0 To UBound(Txt)

            If InStr(1, types(i), "n") Then rs(i) = IIf(Txt(i) = "", 0, Txt(i))
            If InStr(1, types(i), "r") Then rs(i) = IIf(Txt(i) = "", 0, Txt(i))
            If InStr(1, types(i), "d") Then
                If Txt(i) <> "" Then rs(i) = Format(Txt(i), "yyyy/M/d") Else rs(i) = Null
                'If TXT(i) <> "" Then rs(i) = TXT(i)
            End If

            If InStr(1, types(i), "s") Then
                If IsMissing(update) Then
                    If Txt(i) <> "" Then rs(i) = IIf(IsNull(Txt(i)), Null, Txt(i))
                Else
                    rs(i) = IIf(IsNull(Txt(i)) Or Txt(i) = "", Null, Txt(i))

                End If
            End If

            Debug.Print rs(i).Name, rs(i).value, types(i)
        Next i

    End If

    rs.update
    Insert = True
    Exit Function
err_trap:
    Insert = False
    rs.CancelUpdate
    'Resume
    Err.Clear
    Cn.Errors.Clear
End Function

Public Sub fill_combo(My_Combo As DataCombo, _
   My_SQL As String)
    On Error Resume Next
    Dim rs As ADODB.Recordset

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If

    My_Combo.Tag = My_SQL
    Set rs = New ADODB.Recordset
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly
    Debug.Print My_SQL & CHR(13) & rs.RecordCount

    If rs.RecordCount > 0 Then
        'populate the ADO datacombo by setting its properties
        With My_Combo
            '.Text = rs(0)
            Set .RowSource = rs
            .BoundColumn = rs(0).Name
            .ListField = rs(1).Name
            .BoundText = ""
            .text = ""
        End With

    Else
        My_Combo.ReFill
    End If

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

Public Function open_my_connection(Optional changeserver As Boolean = False) As Boolean
    Dim Msg           As String
    Dim StrServerName As String
    Dim IntRes        As Integer

    On Error GoTo ErrTrap
    open_my_connection = False

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        If SystemOptions.SysRegisterState <> DevelopVersion Then
            SetAttr App.path & "\DataFiles\TestData.mdb", vbNormal

            If Dir("D:\bisegypt\AutoBackup", vbDirectory) = "" Then
                CreateFolder "D:\bisegypt\AutoBackup"
            End If

            If Dir("D:\bisegypt\AutoBackup\TestData.mdb", vbNormal) <> "" Then
                Kill "D:\bisegypt\AutoBackup\TestData.mdb"
            End If

            'FileCopy App.Path & "\DataFiles\TestData.mdb", _
             "D:\bisegypt\AutoBackup\TestData.mdb"
        End If
        Cn.Close
        With Cn
            .ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\DataFiles\TestData.mdb;Persist Security Info=False "
            .Open
        End With

        '  UpdateDataBase
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '    If SystemOptions.SysRegisterState = DevelopVersion Then
        '        SystemOptions.SysSQLServerType = LocalServer
        '        SystemOptions.SysSQLServerName = "(LOCAL)"
        '    Else
        '        SystemOptions.SysSQLServerType = Val(GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerType", 0))
        '        SystemOptions.SysSQLServerName = GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", "")
        '    End If
        SystemOptions.SysSQLServerType = val(GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerType", 0))
        If SystemOptions.SysSQLServerName = "" Then SystemOptions.SysSQLServerName = "."
        SystemOptions.SysSQLServerName = GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", "")
        SystemOptions.SysSQLServerTypeTechnical = GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "SysSQLServerTypeTechnical", "0")
     
        'SystemOptions.SysSQLServerName = "Vb-Full"
        If changeserver = True Or SystemOptions.SysSQLServerType = NotSet Or SystemOptions.SysSQLServerName = "" Then
TryConnect:
            Load FrmSQLConData
            FrmSQLConData.show vbModal

            If FrmSQLConData.UserCanceled = True Then
                open_my_connection = False
                Unload FrmSQLConData
                Exit Function
            Else
                Unload FrmSQLConData
            End If
        End If

        Set Cn = New ADODB.Connection

        With Cn
            .CommandTimeout = 0
            .CursorLocation = adUseClient
            .ConnectionTimeout = 30
            '.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
             "Persist Security Info=False;Initial Catalog=SmallAccount;Data Source=(LOCAL)"
            '-------------------------SQL Server 2005
            '        Cn.ConnectionString = "PROVIDER=SQLOLEDB.1;PASSWORD=nour1234nour;" & _
            '        "PERSIST SECURITY INFO=TRUE;USER ID=sa;INITIAL CATALOG=SmallAccount;" & _
            '        "DATA SOURCE=NOURLAPTOP\SQLEXPRESS"
            
            '-------------------------
            '**************************************online*************************************
            If checkonlinedate = True Then '''online
                SystemOptions.SysSQLServerName = onlineservername
                SystemOptions.SysSQLServerDataBaseName = onlineDataBasename
                SystemOptions.SysSQLServerUserId = onlinusername
                SystemOptions.SysSQLServerUserpassword = onlinepassword
                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SystemOptions.SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SystemOptions.SysSQLServerUserId & ";Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName
                GoTo online:
            End If
            '**************************************online*************************************
          
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & "Persist Security Info=False;Initial Catalog=S" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName & ""
            
            If SystemOptions.SysSQLServerType = LocalServer Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & "Persist Security Info=False;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName & ";Port=1433"
        
            ElseIf SystemOptions.SysSQLServerType = RemoteServer Then
                '       .ConnectionString = "Provider=SQLOLEDB.1;Password=salim;Persist Security Info=True; User ID=salim;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName
     
                If SystemOptions.SysSQLServerTypeTechnical = server1 Then
                    .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & "Persist Security Info=False;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName & ";Port=1433"
                Else
                    '        .ConnectionString = "Provider=SQLOLEDB.1;Password=salim;Persist Security Info=True;User ID=salim;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName
                    If SystemOptions.SysSQLServerUserpassword = "" Then SystemOptions.SysSQLServerUserpassword = "Admin@123"
                    If SystemOptions.SysSQLServerUserId = "" Then SystemOptions.SysSQLServerUserId = "sa"
                    .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SystemOptions.SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SystemOptions.SysSQLServerUserId & ";Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName
               
                End If

                '      .ConnectionString = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True; User ID=sa;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & SystemOptions.SysSQLServerName
                'server-t\server_x
            End If
            '.ConnectionString = "Provider=MSDAORA.1;Password=dbo_bytee;User ID=dbo_bytee;Data Source=ORCL;Persist Security Info=True"
online:
            .Open
        End With

    End If

    Dim i As Integer
    'For I = 0 To Cn.Properties.count - 1
    '    Debug.Print I, Cn.Properties(I).name, Cn.Properties(I).Value
    'Next I
    open_my_connection = True
    Exit Function
ErrTrap:
    open_my_connection = False

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        Msg = "Cant Locate File ms.d"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ ⁄œ„ «·⁄»À »„·›«  «·»—‰«„Ã"
        Msg = Msg & CHR(13) & Err.Description
        Msg = Msg & CHR(13) & Err.Number
        Msg = Msg & CHR(13) & Err.Source
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical, App.Title
    Else
        Msg = "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï «·√ ’«· »«·”Ì—›— .."
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ ⁄„· «·‘»ﬂ…"
        Msg = Msg & CHR(13) & Err.Description
        Msg = Msg & CHR(13) & Err.Number
        Msg = Msg & CHR(13) & Err.Source
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical, App.Title
        Msg = "Â·  —Ìœ ≈⁄«œ…  Œ’Ì’ ÃÂ«“ «·”Ì—›— ··»—‰«„Ã ..ø"
        IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

        If IntRes = vbYes Then
            Resume TryConnect
        End If
    
    End If

    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical, App.Title
End Function



Public Function SavePictureToDB(PictControl As Object, _
                                rs As Object, _
                                FieldName As String, Optional sDir As String) As Boolean
    Dim oPict As StdPicture

   ' Dim sDir As String
    Dim sTempFile As String
    Dim iFileNum As Integer
    Dim lFileLength As Long

    Dim abBytes() As Byte
    Dim iCtr As Integer

   On Error GoTo ErrorHandler

    If Not TypeOf rs Is ADODB.Recordset Then Exit Function
    Set oPict = PictControl.Picture

    If oPict Is Nothing Then Exit Function

    'Save picture to temp file
    'sDir = App.Path
    If sDir = "" Then sDir = App.path & "" ' "C:\"
    sTempFile = sDir & "0X2341KLZX.dat"

    If Dir(sTempFile) <> "" Then
        Kill sTempFile
    End If

    SavePicture oPict, sTempFile

    'read file contents to byte array
    iFileNum = FreeFile
    Open sTempFile For Binary Access Read As #iFileNum
    lFileLength = LOF(iFileNum)
    ReDim abBytes(lFileLength)
    Get #iFileNum, , abBytes()
    'put byte array contents into db field
    rs.Fields(FieldName).AppendChunk abBytes()
    Close #iFileNum
    'Don't return false if file can't be deleted
    On Error Resume Next
    Kill sTempFile
    SavePictureToDB = True
ErrorHandler:
End Function



Public Function LoadPictureFromDB(PictControl As Object, _
                                  rs As Object, _
                                  FieldName As String, _
                                  Optional ByRef StrFileName As String) As Boolean

    Dim oPict       As StdPicture
    '   Dim sDir        As String
    Dim sTempFile   As String
    Dim iFileNum    As Integer
    Dim lFileLength As Long
    Dim abBytes()   As Byte
    Dim iCtr        As Integer

    On Error GoTo ErrorHandler

    If Not TypeOf rs Is ADODB.Recordset Then
        Exit Function
    End If
    
    '**********************
    Dim TempPath    As String
    Dim TempFile    As String
    Dim slength     As Long
    Dim lastfour    As Long
    Dim tmpFileName As String
    TempPath = Space(255)
    slength = GetTempPath(255, TempPath)
    TempPath = left(TempPath, slength)
  
    TempFile = Space(255)
    lastfour = GetTempFileName(TempPath, "CR", 0, TempFile)
    TempFile = left(TempFile, InStr(TempFile, vbNullChar) - 1)
    tmpFileName = TempFile
    '***************************

    'sDir = GetTempDir
    ' If sDir = "" Then sDir = "C:\"
    
    sTempFile = tmpFileName 'sDir & "0X2341KLZX.dat"

    If Len(Dir$(sTempFile)) > 0 Then
        On Error Resume Next
        Kill sTempFile
    End If

    iFileNum = FreeFile
    Open sTempFile For Binary As #iFileNum
    lFileLength = LenB(rs(FieldName) & "")
    If lFileLength = 0 Then
        Exit Function
    End If
    abBytes = rs(FieldName).GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum

    If Not PictControl Is Nothing Then
        PictControl.Picture = LoadPicture(sTempFile)
        Kill sTempFile
        StrFileName = sTempFile
    Else
        StrFileName = sTempFile
    End If

    LoadPictureFromDB = True
ErrorHandler:

End Function


Public Function LoadPictureFromDBdd(PictControl As Object, _
                                  rs As Object, _
                                  FieldName As String, _
                                  Optional ByRef StrFileName As String) As Boolean

    Dim oPict       As StdPicture
    '   Dim sDir        As String
    Dim sTempFile   As String
    Dim iFileNum    As Integer
    Dim lFileLength As Long
    Dim abBytes()   As Byte
    Dim iCtr        As Integer

    On Error GoTo ErrorHandler

    If Not TypeOf rs Is ADODB.Recordset Then
        Exit Function
    End If
    
    '**********************
    Dim sfo As New FileSystemObject
    Dim mrs As New ADODB.Recordset
    mrs.Open "SELECT NEWID() id ", Cn, adOpenForwardOnly, adLockReadOnly
    Dim ID As String
    ID = mrs!ID
    mrs.Close
    Set mrs = Nothing
     
    Dim FullDirName As String
    FullDirName = sfo.BuildPath(sfo.GetSpecialFolder(2), ID)
    sfo.CreateFolder FullDirName
    StrFileName = sfo.BuildPath(FullDirName, "img.bmp")
    
    '    iFileNum = FreeFile
    '    Open sTempFile For Binary As #iFileNum
    '    lFileLength = LenB(rs(FieldName) & "")
    '    If lFileLength = 0 Then
    '        Exit Function
    '    End If
    '    abBytes = rs(FieldName).GetChunk(lFileLength)
    '    Put #iFileNum, , abBytes()
    '    Close #iFileNum
    Dim strStream As ADODB.Stream
    Set strStream = New ADODB.Stream
    strStream.type = adTypeBinary
    strStream.Open
    strStream.Write rs(FieldName)
    strStream.SaveToFile StrFileName, adSaveCreateOverWrite
    strStream.Close
    Set strStream = Nothing
    LoadPictureFromDBdd = True
   
    If Not PictControl Is Nothing Then
        PictControl.Picture = LoadPicture(StrFileName)
    End If

    LoadPictureFromDBdd = True
    Exit Function
ErrorHandler:
    LoadPictureFromDBdd = False
End Function
'
'Public Function LoadPictureFromDBToTemp(rs As Object, _
'                                        FieldName As String, _
'                                        Optional ByRef StrFileName As String) As String
'
'    Dim oPict       As StdPicture
'    Dim sDir        As String
'    Dim sTempFile   As String
'    Dim iFileNum    As Integer
'    Dim lFileLength As Long
'    Dim abBytes()   As Byte
'    Dim iCtr        As Integer
'
'    On Error GoTo ErrorHandler
'
'    If Not TypeOf rs Is ADODB.Recordset Then
'        Exit Function
'    End If
'
'    '*****************
'    Dim TempPath    As String
'    Dim TempFile    As String
'    Dim slength     As Long
'    Dim lastfour    As Long
'    Dim tmpFileName As String
'    TempPath = Space(255)
'    slength = GetTempPath(255, TempPath)
'    TempPath = left(TempPath, slength)
'
'    TempFile = Space(255)
'    lastfour = GetTempFileName(TempPath, "PIC", 0, TempFile)
'    TempFile = left(TempFile, InStr(TempFile, vbNullChar) - 1)
'    tmpFileName = TempFile
'    '*******************
'    sTempFile = tmpFileName 'sDir & "0X2341KLZX.dat"
'
'    If Len(Dir$(sTempFile)) > 0 Then
'        On Error Resume Next
'        Kill sTempFile
'    End If
'
'    iFileNum = FreeFile
'    Open sTempFile For Binary As #iFileNum
'    lFileLength = LenB(rs(FieldName) & "")
'    If lFileLength = 0 Then
'        Exit Function
'    End If
'    abBytes = rs(FieldName).GetChunk(lFileLength)
'    Put #iFileNum, , abBytes()
'    Close #iFileNum
'
'    StrFileName = sTempFile
'
'
'    LoadPictureFromDBToTemp = StrFileName
'ErrorHandler:
'End Function

Public Function ChequeBoxOperations(Optional NoteID As Long) As Boolean
    If SystemOptions.updatecashvchrifdeposite = True Then
        ChequeBoxOperations = True
        Exit Function
    End If
    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String
    'On Error GoTo ErrTrap

    If SystemOptions.ChequeBox = False Then
        ChequeBoxOperations = True
        Exit Function
    End If
   
    ChequeBoxOperations = True
 
    If SystemOptions.SysDataBaseType = AccessDataBase Then
   
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  * from TblChecqueBoxContent "
    
        StrSQL = StrSQL + " where  (Deposited=1 or Collected=1) and  NOTEID =" & NoteID
  
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        ChequeBoxOperations = False
        Exit Function
    Else
   
        ChequeBoxOperations = True
 
        Exit Function
    End If

End Function
 
Public Function ChequeBoxOperations1(Optional NoteID As Long) As Boolean

    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String
    'On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = False Then
        ChequeBoxOperations1 = True
        Exit Function
    End If
   
    ChequeBoxOperations1 = True
 
    If SystemOptions.SysDataBaseType = AccessDataBase Then
   
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  * from TblChecqueBoxContent1 "
    
        StrSQL = StrSQL + " where  (Payed=1) and  NOTEID =" & NoteID
  
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        ChequeBoxOperations1 = False
        Exit Function
    Else
   
        ChequeBoxOperations1 = True
 
        Exit Function
    End If

End Function

Public Function ChequeBoxCollect(Optional NoteID As Long, _
                                 Optional Returntransaction As Long = 0) As Boolean

    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String
    'On Error GoTo ErrTrap

    If SystemOptions.ChequeBox = False Then
        ChequeBoxCollect = True
        Exit Function
    End If
   
    ChequeBoxCollect = True
 
    If SystemOptions.SysDataBaseType = AccessDataBase Then
   
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  * from TblChecqueBoxContent "
    
        StrSQL = StrSQL + " where  ( Collected=1) and  NOTEID =" & NoteID
        If Returntransaction <> 0 Then
            StrSQL = StrSQL + " and  NOTEID =" & Returntransaction
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        ChequeBoxCollect = False
        Exit Function
    Else
   
        ChequeBoxCollect = True
 
        Exit Function
    End If

End Function
Public Function CheckBoxmaxVaue(Account_code As String, _
                                currentvalue As Double, _
                                Optional ByRef maxvalue As Double) As Boolean

    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String

    StrSQL = "SELECT     TOP 100 PERCENT boxValue From dbo.TblBoxesData Where (Account_Code ='" & Account_code & "')"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'Dim Priod As Integer
   
    If rs.RecordCount > 0 Then
 
        maxvalue = IIf(IsNull(rs("boxValue").value), 0, rs("boxValue").value)
         
        If maxvalue = 0 Then
        
            CheckBoxmaxVaue = True
            Exit Function
        Else
        
            '    Dim acc As String
            Dim Balance As String
            'acc = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", boxId)
            WriteCustomerBalPublic Account_code, Balance ',  balanceString
     
            If Balance + val(currentvalue) <= maxvalue Then
                
                CheckBoxmaxVaue = True
                Exit Function
            Else
                
                CheckBoxmaxVaue = False
                Exit Function
            End If
        
        End If
  
    Else
   
    End If
   
End Function
Public Function CheckBoxAccountTimes(BoxID As Integer, _
                                     transactiondata As Date, _
                                     Optional ByRef Priod As Integer) As Boolean
    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String
    'On Error GoTo ErrTrap
   
    StrSQL = "SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.TblBoxesData.Priod"
    StrSQL = StrSQL + "  FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID RIGHT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblBoxesData ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.TblBoxesData.Account_Code"
    StrSQL = StrSQL + " Where (dbo.Notes.NoteType = 50) And (dbo.TblBoxesData.boxId = " & BoxID & ")"
    StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate DESC"
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'Dim Priod As Integer
    Dim LastDate  As Date
    Dim datediffe As Integer
    If rs.RecordCount > 0 Then
 
        Priod = IIf(IsNull(rs("Priod").value), 0, rs("Priod").value)
        LastDate = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
         
        If Priod = 0 Then
            CheckBoxAccountTimes = True
            Exit Function
        Else
            datediffe = DateDiff("d", LastDate, transactiondata)
            If datediffe <= Priod Then
                
                CheckBoxAccountTimes = True
                Exit Function
            Else
                
                CheckBoxAccountTimes = False
                Exit Function
                        
            End If
        
        End If
         
        Exit Function
    Else
   
        CheckBoxAccountTimes = True
        Exit Function
    End If

End Function
Public Function CheckBoxAccount(LngBoxID As Long, _
                                DblOutCash As Double, _
                                D_Date As Date, _
                                Optional ShowMsg As Boolean = True, _
                                Optional ByRef DblExistValue As Double, _
                                Optional LngExceptNoteID As Long) As Boolean

    Dim rs              As ADODB.Recordset
    Dim StrSQL          As String
    Dim DblExistAccount As Double
    Dim Msg             As String
    Dim StrBoxName      As String
    'On Error GoTo ErrTrap
    CheckBoxAccount = True
    Exit Function
    If SystemOptions.SysAllowBoxNegative = True Then
        CheckBoxAccount = True
        Exit Function
    End If
   
    CheckBoxAccount = True

    '------------------------------------------------------------
    'StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & _
    '" FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & _
    '"QryBoxesCredit.BoxID "
    'StrSQL = StrSQL + " Where TblBoxesData.BoxID=" & LngBoxID & ""
    '--------------------------------------------------------------
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT QryBoxBalance.BoxID, QryBoxBalance.BoxName," & "Sum(QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From QryBoxBalance "
        StrSQL = StrSQL + " Where (QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " AND QryBoxBalance.NoteDate <= #" & Format(D_Date, "MM/dd/yyyy") & "#"

        If LngExceptNoteID <> 0 Then
            StrSQL = StrSQL + " AND NOTEID <>" & LngExceptNoteID
        End If

        StrSQL = StrSQL + " GROUP BY QryBoxBalance.BoxID, QryBoxBalance.BoxName;"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.QryBoxBalance.BoxID, dbo.QryBoxBalance.BoxName," & "Sum(dbo.QryBoxBalance.Note_Value  * TransDir) AS BoxAccount " & "From dbo.QryBoxBalance() "
        StrSQL = StrSQL + " Where (dbo.QryBoxBalance.BoxID=" & LngBoxID & ") "
        StrSQL = StrSQL + " AND dbo.QryBoxBalance.NoteDate <=" & SQLDate(D_Date, True) & ""

        If LngExceptNoteID <> 0 Then
            StrSQL = StrSQL + " AND NOTEID <>" & LngExceptNoteID
        End If

        StrSQL = StrSQL + " GROUP BY dbo.QryBoxBalance.BoxID, dbo.QryBoxBalance.BoxName;"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        'DblExistAccount = IIf(IsNull(Rs("BoxCredit").Value), 0, Rs("BoxCredit").Value)
        If IsNull(rs("BoxAccount").value) Then
            DblExistAccount = 0
        Else
            DblExistAccount = Round(rs("BoxAccount").value, SystemOptions.SysDefCurrencyForamt)
        End If

        StrBoxName = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
    
        If Not IsMissing(DblExistValue) Then
            DblExistValue = DblExistAccount
        End If
  
        Dim FirstPeriod As Date
        Dim AccountCode As String
        getFirstPeriodDateInthisYear FirstPeriod
        AccountCode = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", LngBoxID)
        DblExistAccount = GetActualAccountBalance(AccountCode, branch_id, FirstPeriod, Date)

        If DblOutCash > DblExistAccount Then
            If ShowMsg = True Then
                Msg = "⁄›Ê« ·«Ì„ﬂ‰ «·”„«Õ »≈ „«„ «·⁄„·Ì…"
                Msg = Msg & CHR(13) & "ÕÌÀ «‰ «·—’Ìœ «·Õ«·Ï ›Ï «·Œ“‰… «·„Õœœ…"
                Msg = Msg & CHR(13) & "«ﬁ· „‰ «·„»·€ «·„—«œ "
                Msg = Msg & CHR(13) & ""
                Msg = Msg & CHR(13) & "«·—’Ìœ «·Õ«·Ï ›Ï «·Œ“‰…  "

                If DblExistAccount < 0 Then
                    Msg = Msg & CHR(13) & "œ«∆‰"
                End If

                Msg = Msg & CHR(13) & StrBoxName & "=" & DblExistAccount
                Msg = Msg & CHR(13) & "«·„»·€ «·„—«œ Œ—ÊÃÂ = " & DblOutCash
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If

            Screen.MousePointer = vbDefault
            CheckBoxAccount = False
        End If

    Else

        If ShowMsg = True Then
            Msg = "·«ÌÊÃœ «Ï —’Ìœ ›Ï «·Œ“‰… «·„Õœœ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If

        Screen.MousePointer = vbDefault
        CheckBoxAccount = False
    End If

    Exit Function

ErrTrap:
    Resume
    Msg = "⁄›Ê«  ⁄œ— Õ”«» «·—’Ìœ «·Õ«·Ï ›Ï «·Œ“‰…... !!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    CheckBoxAccount = False
End Function

Public Sub ShowBoxesAccouns()
    Dim rs          As ADODB.Recordset
    Dim StrSQL      As String
    Dim Msg         As String
    Dim i           As Integer
    Dim FirstPeriod As Date
    Dim Balance     As Double
    'On Error GoTo ErrTrap
    'StrSQL = "SELECT * from TblBoxesData where type=0 "
    StrSQL = "SELECT * from TblBoxesData  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Load FrmBoxesAccounts

        With FrmBoxesAccounts.FgBoxes
            .rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                Else
                    .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNamee").value), "", rs("BoxNamee").value)
                End If
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
      
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, 0, FirstPeriod, Date)
            
                '        .TextMatrix(i, .ColIndex("BoxCredit")) = (get_balanceFromGl(rs("Account_Code").value))
                .TextMatrix(i, .ColIndex("BoxCredit")) = Abs(Balance) 'GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "„œÌ‰"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "œ«∆‰"
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                Else

                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Debit"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Credit"
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Exit Sub

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where TblBoxesData.BoxID <>1"
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblBoxesData.BoxID,dbo.TblBoxesData.BoxName, QryBoxesCredit.BoxCredit" & " FROM dbo.TblBoxesData INNER JOIN " & "dbo.QryBoxesCredit() QryBoxesCredit ON dbo.TblBoxesData.BoxID = QryBoxesCredit.BoxID"

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <>1"
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Load FrmBoxesAccounts

        With FrmBoxesAccounts.FgBoxes
            .rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)

                If Not IsNull(rs("BoxCredit").value) Then
                    .TextMatrix(i, .ColIndex("BoxCredit")) = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
                Else
                    .TextMatrix(i, .ColIndex("BoxCredit")) = 0
                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

        FrmBoxesAccounts.show
        FrmBoxesAccounts.ZOrder 0
    Else
        Msg = "·«ÌÊÃœ «Ï Œ“‰ „”Ã·… ›Ï «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

End Sub

Public Function GetBoxAccount(LngBoxID As Long) As Double
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg    As String
    Dim i      As Integer
    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "
        StrSQL = StrSQL + " Where QryBoxesCredit.BoxID=" & LngBoxID & ""
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN dbo.QryBoxesCredit()QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "
        StrSQL = StrSQL + " Where QryBoxesCredit.BoxID=" & LngBoxID & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        If IsNull(rs("BoxCredit").value) Then
            GetBoxAccount = 0
        Else
            GetBoxAccount = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
        End If

    Else
        GetBoxAccount = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    GetBoxAccount = 0
End Function
Public Function GetItemPrice(LngItemID As Long, _
                             Optional DblQty As Double = 1, _
                             Optional LngUnitID As Long, _
                             Optional Purchase As Long = 0, _
                             Optional bycustomerpolitical As Integer) As Double
    
    Dim DblRes  As Double
    Dim StrSQL  As String
    Dim RsPrice As ADODB.Recordset
    Dim RsTemp  As ADODB.Recordset
    Dim NewGrid As New ClsGrid
    On Error GoTo ErrTrap

    If DblQty = 0 Then
        DblRes = 0
    End If
                
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Min(ItemsPrice.From) AS MinQty , Max(ItemsPrice.To) AS " & "MaxQty From ItemsPrice GROUP BY ItemsPrice.Item_ID " & "Having ItemsPrice.Item_ID= " & LngItemID
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            
            
            
            StrSQL = "SELECT Min(ItemsPrice.[From]) AS MinQty , Max(ItemsPrice.[To]) AS " & _
            "MaxQty From ItemsPrice " & _
            "GROUP BY ItemsPrice.Item_ID " & _
            "Having ItemsPrice.Item_ID= " & LngItemID
    
      ' StrSQL = "Select * From TblItemsUnits where ItemID=" & LngItemID & " and UnitID =" & LngUnitID

    End If

    Set RsPrice = New ADODB.Recordset
    RsPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsPrice.EOF Then
        Set RsPrice = New ADODB.Recordset
        If PPointID <> 0 Then
            StrSQL = "SELECT * FROM Tblposdata where BoxId = " & PPointID & " And IsNull(PriceId ,0) <> 0 "
            RsPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not RsPrice.EOF Then
                StrSQL = "SELECT Price" & Trim(RsPrice!PriceID & "") & " as UnitSalesPrice,Price" & Trim(RsPrice!PriceID & "") & " as    UnitWholeSalePrice FROM TblSalesPrices where BoxId = " & PPointID & " and ItemID= " & LngItemID
                RsPrice.Close
                Set RsPrice = New ADODB.Recordset
                RsPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If Not RsPrice.EOF Then
                    DblRes = val(RsPrice!UnitSalesPrice & "")
                End If
                If DblRes <> 0 Then
                    GetItemPrice = DblRes
                    Exit Function
                End If
            End If
            RsPrice.Close
            StrSQL = "Select * From TblItemsUnits where ItemID=" & LngItemID & " and UnitID =" & LngUnitID
        End If
        StrSQL = "Select * From TblItemsUnits where ItemID=" & LngItemID & " and UnitID =" & LngUnitID
        RsPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If Not (RsPrice.EOF Or RsPrice.BOF) Then
        If DblQty >= 1 Then
            '           If 1 = 1 Then
            '    If DblQty >= 1 And DblQty < Val(RsPrice("UnitSalesPrice").value) Then
            '        StrSQL = "Select * From TblItems where ItemID=" & LngItemID

            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then

                '            DblRes = IIf(IsNull(RsTemp("SallingPrice").Value), 0, RsTemp("SallingPrice").Value)
                If Purchase = 1 Then
                    DblRes = IIf(IsNull(RsTemp("UnitPurPrice").value), 0, RsTemp("UnitPurPrice").value)
                Else
                
                    If SystemOptions.AllowLastPrice = True Then
                        GetItemPrice = GetLastPrice(LngItemID, LngUnitID)
                        If GetItemPrice <> 0 Then
                            Exit Function
                        End If
                    Else
                    
                    End If
                    If bycustomerpolitical = 0 Then
                        DblRes = IIf(IsNull(RsTemp("UnitSalesPrice").value), 0, RsTemp("UnitSalesPrice").value)
                    Else
                        DblRes = IIf(IsNull(RsTemp("UnitWholeSalePrice").value), 0, RsTemp("UnitWholeSalePrice").value)
                    
                    End If
                    
                End If
            End If

            RsTemp.Close
            RsPrice.Close
        ElseIf DblQty > val(RsPrice("MaxQty").value) Then
            StrSQL = "select * From ItemsPrice where Item_ID=" & LngItemID
            StrSQL = StrSQL + " and To=" & RsPrice("MaxQty").value
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                DblRes = IIf(IsNull(RsTemp("Price").value), 0, RsTemp("Price").value)
            End If

            RsTemp.Close
            RsPrice.Close
        Else
            StrSQL = "select * From ItemsPrice where Item_ID=" & LngItemID
            StrSQL = StrSQL + " and [From] <=" & DblQty
            StrSQL = StrSQL + " and [To] >=" & DblQty
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                DblRes = IIf(IsNull(RsTemp("Price").value), 0, RsTemp("Price").value)
            End If

            RsTemp.Close
            RsPrice.Close
        End If

    Else
        StrSQL = "select * From TblItems where ItemID=" & LngItemID
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            DblRes = IIf(IsNull(RsTemp("SallingPrice").value), 0, RsTemp("SallingPrice").value)
        End If

        If DblRes = 0 Then
                
                
                StrSQL = "SELECT  TblItemsUnits.UnitSalesPrice FROM TblItemsUnits  WHERE ItemID =" & LngItemID & "  AND DefaultUnit = 1"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    DblRes = IIf(IsNull(RsTemp("UnitSalesPrice").value), 0, RsTemp("UnitSalesPrice").value)
                End If
            
        End If

        RsTemp.Close
    End If

    GetItemPrice = DblRes
    Exit Function
ErrTrap:
    GetItemPrice = 0
End Function

Public Function GetLastItemCode(LngGroupID As Long) As String
    Dim rs      As ADODB.Recordset
    Dim StrSQL  As String
    Dim StrTemp As String
    On Error GoTo ErrTrap

    StrSQL = "SELECT TblItems.ItemID, TblItems.ItemCode, Groups.GroupID "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
    StrSQL = StrSQL + " Where Groups.GroupID =" & LngGroupID & ""
    StrSQL = StrSQL + " Order by  Groups.GroupID,TblItems.ItemID "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
        StrTemp = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
    Else
        StrTemp = ""
    End If

    rs.Close
    Set rs = Nothing
    GetLastItemCode = StrTemp
    Exit Function
ErrTrap:
    GetLastItemCode = ""
End Function

Public Function GetTransIDFromNoteSerial1(NoteSerial1 As String, _
                                          Optional ByRef Transaction_ID As Long, _
                                          Optional ByRef Transaction_Date As Date, _
                                          Optional Transaction_Type As Integer, _
                                          Optional SpecialOffer As Integer)

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Select * From Transactions Where NoteSerial1='" & NoteSerial1 & "'"
    StrSQL = StrSQL + " AND (Transaction_Type=" & Transaction_Type & " )"
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Transaction_ID = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
        Transaction_Date = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
        SpecialOffer = IIf(IsNull(rs("SpecialOffer").value), 0, rs("SpecialOffer").value)
 
    End If
 
End Function

Public Function GetTransNoteSerial1FromID(Transaction_ID As Long, _
                                          Optional ByRef NoteSerial1 As String)

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Select * From Transactions Where Transaction_ID=" & Transaction_ID
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        NoteSerial1 = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    
    End If
 
End Function

Public Function GetTransNoteSerial1TransactionSerail(Transaction_serial As String, _
                                                     Optional ByRef NoteSerial1 As String)

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Select * From Transactions Where  Transaction_Type=21 and Transaction_Serial='" & Transaction_serial & "'"
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        NoteSerial1 = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    
    End If
 
End Function

Public Function GetTransIDSerial(IntType As Integer, _
                                 Optional LngTransID As Long = 0, _
                                 Optional StrTransSerial As String = "", _
                                 Optional IntTransType As String, _
                                 Optional LngCusID As Long = 0, _
                                 Optional InvID As Integer) As String
    
    'IntType=0 --- Get Transaction_ID
    'IntType=1 ---Get TransactionSerial
    'IntTransType Values  See the  Transactions_Type Table in DataBase

    '„·ÕÊŸ… Â«„…
    '⁄‰œ„«  —Ìœ «·”Ì—Ì· „‰ «·œ«·…
    '›ÌÃ» ⁄·Ìﬂ «‰  —”· ··œ«·… «·‹
    'Transaction_ID
    'Transaction_Type
    '·«‰ «·”Ì—Ì· ·ÊÕœÂ ·«Ìﬂ›Ï
    '·«‰Â „„ﬂ‰  ﬂ—«— «·”Ì—Ì«·

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim Temp   As String

    If IntType = 0 Then
        StrSQL = "Select * From Transactions Where Transaction_Serial='" & StrTransSerial & "'"
        StrSQL = StrSQL + " AND (Transaction_Type=2 or Transaction_Type=21)"

        If InvID <> 0 Then
            StrSQL = StrSQL + "  And Transaction_id = '" & InvID & "'"
        End If

        If LngCusID <> 0 Then
            StrSQL = StrSQL + " AND CusID=" & LngCusID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Temp = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            LngCusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
            FrmReturnSalling.DBCboClientName.BoundText = LngCusID
            FrmReturnSalling.TxtTransSerial.text = rs("transaction_serial")
            GetTransIDSerial = rs("transaction_serial")

        End If
    
    ElseIf IntType = 1 Then
        StrSQL = "Select * From Transactions Where Transaction_ID=" & InvID & ""
        StrSQL = StrSQL + ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Temp = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
            LngCusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
            FrmReturnpurchases.DBCboClientName.BoundText = LngCusID
            FrmReturnpurchases.TxtTransSerial.text = Temp
            GetTransIDSerial = Temp

        End If
    End If

End Function

Public Function GetManIDSerial(IntType As Integer, _
                               Optional LngTransID As Long = 0, _
                               Optional StrTransSerial As String = "", _
                               Optional IntTransType As Integer, _
                               Optional LngCusID As Long = 0, _
                               Optional StrCashCusName As String, _
                               Optional LngStoreID As Long = 0, _
                               Optional LngItemID As Long = 0, _
                               Optional StrItemSerial As String = "", _
                               Optional Quantity As Long = 0) As String
    
    'IntType=0 --- Get MaintananceID
    'IntType=1 ---Get ReciptNumber
    'IntTransType Values  See the  Transactions_Type Table in DataBase

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim Temp   As String

    If IntType = 0 Then
        StrSQL = "Select * From TblMaintenece Where ReciptNumber='" & StrTransSerial & "'"
        StrSQL = StrSQL + " AND ManOperationTypeID=" & IntTransType & ""

        If LngCusID <> 0 Then
            StrSQL = StrSQL + " AND CusID=" & LngCusID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Temp = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
            LngCusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        End If

        If Not IsMissing(StrCashCusName) Then
            StrCashCusName = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
        End If

    ElseIf IntType = 1 Then
        StrSQL = "Select * From TblMainteneceNew Where ReciptNumber = N'" & LngTransID & "'"
        StrSQL = StrSQL + ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Temp = IIf(IsNull(rs("ReciptNumber").value), "", rs("ReciptNumber").value)
            LngCusID = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
            LngStoreID = IIf(IsNull(rs("StoreID").value), 0, rs("StoreID").value)
         
            StrItemSerial = IIf(IsNull(rs("Itemserial").value), "", rs("Itemserial").value)
            LngItemID = IIf(IsNull(rs("ItemID").value), 0, rs("ItemID").value)
            Quantity = IIf(IsNull(rs("Quantity").value), 0, rs("Quantity").value)
        
        End If

        If Not IsMissing(StrCashCusName) Then
            StrCashCusName = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
        End If
    End If

    GetManIDSerial = Temp
End Function

Public Function GetDealerType(LngDealerID As Long) As Integer
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From TblCustemers Where CusID=" & LngDealerID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        '1-⁄„Ì·
        '2-„Ê—œ
        '3-„ ⁄«„·Ê‰ «Ê „ ⁄·ﬁ« 
        GetDealerType = IIf(IsNull(rs("Type").value), -1, rs("Type").value)
    Else
        GetDealerType = -1
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetDealerID(StrDealerName As String) As Long
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From TblCustemers Where CusName='" & StrDealerName & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        '1-⁄„Ì·
        '2-„Ê—œ
        GetDealerID = IIf(IsNull(rs("CusID").value), -1, rs("CusID").value)
    Else
        GetDealerID = -1
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function Check_CheckNum(StrCheckNum As String, _
                               LngTransID As Long, _
                               StrMode As String, _
                               IntTransType As Integer) As Boolean

    Dim StrSQL As String
    Dim rs     As ADODB.Recordset
    Dim Msg    As String

    'IntTransType=0 'Transactions
    'IntTransType=1 Mantainace

    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, BanksData.BankName, Notes.ChqueNum, Notes.DueDate," & "Transactions.Transaction_Serial, Transactions.Transaction_Date, TblNotesTypes.NotesTypeName," & "TransactionTypes.TransactionTypeName, TblCustemers.CusName, TblMaintenece.MaintananceID," & "Notes.BankID "
    StrSQL = StrSQL + " FROM TransactionTypes RIGHT JOIN (Transactions RIGHT JOIN (TblNotesTypes " & "INNER JOIN (TblMaintenece RIGHT JOIN (TblCustemers RIGHT JOIN (BanksData RIGHT JOIN Notes " & "ON BanksData.BankID = Notes.BankID) ON TblCustemers.CusID = Notes.CusID) ON " & "TblMaintenece.MaintananceID = Notes.MaintananceID) ON TblNotesTypes.NotesType = Notes.NoteType)" & " ON Transactions.Transaction_ID = Notes.Transaction_ID) ON TransactionTypes.Transaction_Type =" & "Transactions.Transaction_Type "
    StrSQL = StrSQL + " WHERE(Notes.NoteType=2 Or Notes.NoteType=13) "

    If StrMode = "N" Then
        StrSQL = StrSQL + " And Notes.ChqueNum='" & StrCheckNum & "'"
    ElseIf StrMode = "E" Then
        StrSQL = StrSQL + " And Notes.ChqueNum='" & StrCheckNum & "'"

        If IntTransType = 0 Then
            StrSQL = StrSQL + " AND Notes.Transaction_ID <> " & LngTransID & ""
        Else
            StrSQL = StrSQL + " AND Notes.MaintananceID <> " & LngTransID & ""
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Check_CheckNum = True
    Else
        Msg = "⁄›Ê« ÌÊÃœ ‘Ìﬂ „”Ã· ›Ï «·»—‰«„Ã »‰›” «·—ﬁ„"
        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·‘Ìﬂ ›Ï «·»—‰«„Ã"
        Msg = Msg & CHR(13) & ""
        Msg = Msg & CHR(13) & "»Ì«‰«  «·‘Ìﬂ «·„”Ã· ”«»ﬁ« ›Ï «·»—‰«„Ã:-"
        Msg = Msg & CHR(13) & " «—ÌŒ  Õ—Ì— «·‘Ìﬂ:" & IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
        Msg = Msg & CHR(13) & "«”„ «·⁄„Ì·:" & IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
        Msg = Msg & CHR(13) & "«”„ «·»‰ﬂ:" & IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
        Msg = Msg & CHR(13) & " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ:" & IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Check_CheckNum = False
    End If

End Function

Public Function GetLastItemTrans(LngItemID As Long, _
                                 Optional IntTransType As Integer = 0, _
                                 Optional StoreID As Double, _
                                 Optional CusID As Long = 0) As LastItemTransInfo
    
11:
    Dim StrSQL        As String
    Dim Acmd          As ADODB.Command
    Dim XParTransType As ADODB.Parameter
    Dim XParItemID    As ADODB.Parameter
    Dim rs            As ADODB.Recordset
    Dim TempRes       As LastItemTransInfo
    Dim mIsEnter      As Boolean
    ' „ «· ÕÊÌ· ≈·Ï
    'SQL Server
    If SystemOptions.SysDataBaseType = AccessDataBase Then

        StrSQL = "SELECT QryLastItemsPurPrice.LastPurTrans, QryLastItemsPurPrice.Transaction_ID," & "QryLastItemsPurPrice.Transaction_Serial, QryLastItemsPurPrice.Transaction_Date, " & "QryLastItemsPurPrice.ItemSerial, QryLastItemsPurPrice.Price, QryLastItemsPurPrice.ItemName," & "QryLastItemsPurPrice.ItemCode, QryLastItemsPurPrice.Item_ID, TblCustemers.CusName," & "Sum(Transaction_Details.Quantity) AS Qty "
        StrSQL = StrSQL + " FROM (TblCustemers INNER JOIN (QryLastItemsPurPrice INNER JOIN Transactions" & " ON QryLastItemsPurPrice.Transaction_ID = Transactions.Transaction_ID) ON TblCustemers.CusID = " & "Transactions.CusID) INNER JOIN Transaction_Details ON Transactions.Transaction_ID = " & "Transaction_Details.Transaction_ID "
        StrSQL = StrSQL + " GROUP BY QryLastItemsPurPrice.LastPurTrans, QryLastItemsPurPrice.Transaction_ID," & "QryLastItemsPurPrice.Transaction_Serial, QryLastItemsPurPrice.Transaction_Date," & "QryLastItemsPurPrice.ItemSerial, QryLastItemsPurPrice.Price, QryLastItemsPurPrice.ItemName," & "QryLastItemsPurPrice.ItemCode, QryLastItemsPurPrice.Item_ID, TblCustemers.CusName "
    
        Set Acmd = New ADODB.Command
        Set Acmd.ActiveConnection = Cn
        Acmd.CommandType = adCmdUnknown
        Acmd.CommandText = StrSQL
               
        Set XParTransType = New ADODB.Parameter
        XParTransType.type = adInteger
        XParTransType.Name = "X"
        XParTransType.Direction = adParamInput
        XParTransType.value = IntTransType
        Acmd.Parameters.Append XParTransType
    
        Set XParItemID = New ADODB.Parameter
        XParItemID.type = adInteger
        XParItemID.Name = "Y"
        XParItemID.Direction = adParamInput
        XParItemID.value = LngItemID
        Acmd.Parameters.Append XParItemID
    
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenStatic
        rs.LockType = adLockReadOnly
        Set rs = Acmd.Execute()
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        'Stop
        StrSQL = "SELECT QryLastPurItemsTrans.LastPurTrans, dbo.Transactions.NOTESERIAL1, dbo.Transactions.Transaction_ID," & "dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, "
        StrSQL = StrSQL + "dbo.Transaction_Details.Price,dbo.Transaction_Details.ShowPrice," & "dbo.TblItems.ItemName, dbo.TblItems.ItemCode, "
        StrSQL = StrSQL + " QryLastPurItemsTrans.ItemID, dbo.TblCustemers.CusName "
        StrSQL = StrSQL + ",Sum(dbo.Transaction_Details.showqty) as Qty"
        StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN"
        StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID =" & "dbo.Transaction_Details.Transaction_ID INNER JOIN"
        '---------------in the next line we put the transaction type and the item ID
        StrSQL = StrSQL + " dbo.QryLastPurItemsTrans(" & IntTransType & ", " & LngItemID & ",'1-NOV-9999') QryLastPurItemsTrans INNER JOIN"
        '---------------
        StrSQL = StrSQL + " dbo.TblItems ON QryLastPurItemsTrans.ItemID = dbo.TblItems.ItemID " & "ON dbo.Transactions.Transaction_ID = QryLastPurItemsTrans.LastPurTrans  "
        StrSQL = StrSQL + " LEFT OUTER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
        StrSQL = StrSQL + " Where  dbo.Transaction_Details.Item_ID=" & LngItemID & ""
        StrSQL = StrSQL + " Group By QryLastPurItemsTrans.LastPurTrans, dbo.Transactions.Transaction_ID,"
        StrSQL = StrSQL + " dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date,"
        StrSQL = StrSQL + " dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowPrice,dbo.Transactions.NOTESERIAL1,"
        StrSQL = StrSQL + " dbo.TblItems.ItemName, dbo.TblItems.ItemCode,"
        StrSQL = StrSQL + " QryLastPurItemsTrans.ItemID , dbo.TblCustemers.CusName"
        
        'newwwwwwwww
        StrSQL = " SELECT     TOP 1 PERCENT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.Transaction_Details.showPrice, "
        StrSQL = StrSQL + "                       dbo.Transaction_Details.ShowQty, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.ID, dbo.Transactions.NoteSerial1,"
        StrSQL = StrSQL + "                       dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee"
        StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL + "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
        StrSQL = StrSQL + "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
                      
        StrSQL = StrSQL + "  Where (dbo.Transactions.Transaction_Type = " & IntTransType & ") "
        StrSQL = StrSQL + "   And (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
        If StoreID <> 0 Then
            StrSQL = StrSQL + "   And (dbo.Transactions.StoreId = " & StoreID & ")"
        End If
        If CusID <> 0 Then
            StrSQL = StrSQL + "   And (dbo.Transactions.CusID = " & CusID & ")"
        End If

        StrSQL = StrSQL + "  ORDER BY dbo.Transactions.Transaction_Date DESC, dbo.Transactions.Transaction_ID DESC, dbo.Transaction_Details.ID DESC"

        Set rs = New ADODB.Recordset
        
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If

    If rs.BOF And mIsEnter = False And CusID <> 0 Then
        CusID = 0
        mIsEnter = True
        GoTo 11
    End If

    'Rs.Open Acmd, , adOpenStatic, adLockReadOnly
    If Not (rs.BOF Or rs.EOF) Then
        TempRes.Transactionid = IIf(IsNull(rs("Transaction_ID").value), 0, rs("Transaction_ID").value)
        TempRes.TransactionSerial = IIf(IsNull(rs("NOTESERIAL1").value), "", rs("NOTESERIAL1").value)

        If Not IsNull(rs("Transaction_Date").value) Then
            TempRes.TransactionDate = DisplayDate(rs("Transaction_Date").value)
        Else
            TempRes.TransactionDate = DisplayDate(Date)
        End If

        TempRes.SngItemPrice = IIf(IsNull(rs("ShowPrice").value), 0, rs("ShowPrice").value)
        TempRes.SngItemQty = IIf(IsNull(rs("ShowQty").value), 0, rs("ShowQty").value)
        If SystemOptions.UserInterface = ArabicInterface Then
            TempRes.StrCustomerName = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
        Else
            TempRes.StrCustomerName = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
        End If
    End If

    GetLastItemTrans = TempRes
End Function

Public Function GetItemsRsTransactions(Optional LngStoreID As Integer = 0, _
                                       Optional M_Date As Date) As ADODB.Recordset

    Dim AdCmd  As ADODB.Command
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim AdPar  As ADODB.Parameter
    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "Select * From QryItemsInOutTransactions "

        If LngStoreID <> 0 Then
            StrSQL = StrSQL + " Where StoreID=" & LngStoreID & ""
        End If

        Set AdPar = New ADODB.Parameter
        AdPar.Direction = adParamInput
        AdPar.type = adDate
        AdPar.Name = "TransDate"
        AdPar.value = M_Date
        Set AdCmd = New ADODB.Command
        Set AdCmd.ActiveConnection = Cn
        AdCmd.CommandType = adCmdStoredProc
        AdCmd.CommandText = StrSQL
        AdCmd.Parameters.Append AdPar
    
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.LockType = adLockReadOnly
        rs.CursorType = adOpenStatic
        rs.Open AdCmd, , , , adCmdUnknown
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = " SELECT QryItemsInOutTransactions.*"
        StrSQL = StrSQL + " FROM dbo.QryItemsInOutTransactions(" & SQLDate(M_Date, True) & "," & SQLDate(M_Date, True) & ") QryItemsInOutTransactions"

        If LngStoreID <> 0 Then
            StrSQL = StrSQL + " Where StoreID=" & LngStoreID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.LockType = adLockReadOnly
        rs.CursorType = adOpenStatic
        rs.Open StrSQL, Cn
    End If

    Set GetItemsRsTransactions = rs
    Exit Function
ErrTrap:
    Set GetItemsRsTransactions = Nothing
End Function

Public Function GetGovernmentID(StrGovCode As String, _
                                Optional StrGovName As String = "") As Long
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    If Trim(StrGovCode) <> "" Then
        StrSQL = "Select GovernmentID From TblCountriesGovernments Where code='" & Trim(StrGovCode) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetGovernmentID = rs("GovernmentID").value
        End If

        rs.Close
        Set rs = Nothing

    ElseIf Trim(StrGovName) <> "" Then
        StrSQL = "Select GovernmentID From TblCountriesGovernments Where GovernmentName='" & Trim(StrGovName) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetGovernmentID = rs("GovernmentID").value
        End If

        rs.Close
        Set rs = Nothing
    End If

End Function

Public Function GetGovernmentCode(LngGovernmentID As Long) As String
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    If LngGovernmentID <> 0 Then
        StrSQL = "Select code From TblCountriesGovernments Where GovernmentID=" & LngGovernmentID & ""
        Set rs = New ADODB.Recordset

        If Cn.State = adStateClosed Then
            open_my_connection
        End If

        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetGovernmentCode = IIf(IsNull(rs("code").value), "", rs("code").value)
        End If

        rs.Close
        Set rs = Nothing
    End If

End Function

Public Function GetItemID(StrItemCode As String, _
                          Optional StrItemName As String = "") As Long
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    If Trim(StrItemCode) <> "" Then
        If Trim(StrItemCode) = "" Then
            Exit Function
        End If
        StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(StrItemCode) & "' or  barCodeNO='" & Trim(StrItemCode) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetItemID = rs("ItemID").value
        Else
        End If

        rs.Close
        Set rs = Nothing
    ElseIf StrItemName <> "" Then
        StrSQL = "Select ItemID From TblItems Where ItemName='" & Trim(StrItemName) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetItemID = rs("ItemID").value
        Else
        End If

        rs.Close
        Set rs = Nothing
    End If

End Function

Public Function GetItemCode(LngItemID As Long) As String
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    If LngItemID <> 0 Then
        StrSQL = "Select ItemCode  From TblItems Where ItemID=" & LngItemID & ""
        Set rs = New ADODB.Recordset

        If Cn.State = adStateClosed Then
            open_my_connection
        End If

        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            GetItemCode = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        Else
        End If

        rs.Close
        Set rs = Nothing
    End If

End Function

Public Function GetPostTransID() As String
    Dim rs      As ADODB.Recordset
    Dim StrSQL  As String
    Dim i       As Integer
    Dim StrTemp As String
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT QryPostedTransactions.Transaction_ID FROM QryPostedTransactions " & "Order By QryPostedTransactions.Transaction_ID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    StrTemp = ""

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1
            StrTemp = StrTemp & rs("Transaction_ID").value & ","
            rs.MoveNext
        Next i

    End If

    If StrTemp <> "" Then
        StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
    End If

    GetPostTransID = StrTemp
    Exit Function
ErrTrap:
    GetPostTransID = ""
End Function

Public Function GetDebitCreditValues(IntType As Integer, _
                                     m_FromDate As Variant, _
                                     m_ToDate As Variant, _
                                     Optional LngCusID As Long = 0, _
                                     Optional IntWarnDay As Integer) As ADODB.Recordset

    Dim rs             As ADODB.Recordset
    Dim StrSQL         As String
    Dim BolBegain      As Boolean
    Dim StrPostedID    As String
    Dim RsCreditValues As ADODB.Recordset

    StrPostedID = GetPostTransID

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        'Cn.CommandTimeout = 60
        Cn.CommandTimeout = 10000

    End If

    If IntType = 0 Then

        'Debit Values
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT Transaction_ID as TransactionsID, Transaction_Type, TransactionTypeName, Transaction_Date, CusName," & "NoteID, NoteType, Note_Value, DueDate, Transaction_Serial, CusID, NoteDate, NotesTypeName," & "PreRelease, Note_Value-IIF(IsNULL(PreRelease),0,PreRelease) AS RequiredValue " & " From  ( "
            StrSQL = StrSQL + " SELECT PaymentTime.Transaction_ID, PaymentTime.Transaction_Type," & "PaymentTime.TransactionTypeName, PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, " & "PaymentTime.NoteType, PaymentTime.Note_Value, PaymentTime.DueDate, PaymentTime.Transaction_Serial, " & "PaymentTime.CusID, PaymentTime.NoteDate, PaymentTime.NotesTypeName, Sum(QryPostedTransNotes.Note_Value) AS PreRelease" & " FROM PaymentTime LEFT JOIN QryPostedTransNotes ON PaymentTime.Transaction_ID =  " & " QryPostedTransNotes.Transaction_ID"
            StrSQL = StrSQL + " WHERE (((PaymentTime.Transaction_Type)=1 Or (PaymentTime.Transaction_Type)=9) " & "AND ((PaymentTime.NoteID) Not In (select NoteID From InstallMent) And (PaymentTime.NoteID)  "
            StrSQL = StrSQL + " Not In (select NoteID From TblCheckRelease)) "

            If StrPostedID <> "" Then
                'AND ((PaymentTime.Transaction_ID) Not In (1,0)))"
                StrSQL = StrSQL + " and ((PaymentTime.Transaction_ID) Not IN(" & StrPostedID & "))) "
            End If

            StrSQL = StrSQL + " GROUP BY PaymentTime.Transaction_ID, PaymentTime.Transaction_Type, PaymentTime.TransactionTypeName," & "PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, PaymentTime.NoteType, PaymentTime.Note_Value," & "PaymentTime.DueDate, PaymentTime.Transaction_Serial, PaymentTime.CusID, PaymentTime.NoteDate, " & "PaymentTime.NotesTypeName, PaymentTime.NotesTypeName "
            StrSQL = StrSQL + " Union "
            StrSQL = StrSQL + " SELECT   Transaction_ID,Transaction_Type,TransactionTypeName,Transaction_Date, "
            StrSQL = StrSQL + "CusName,NoteID,3 as NoteType,[Value] as Note_Value,DueDate,Transaction_Serial,CustID," & "Transaction_Date,'ﬁ”ÿ „” Õﬁ' as NotesTypeName,Summition as  PreRelease "
            StrSQL = StrSQL + " From QryCust_Qest"
            StrSQL = StrSQL + " WHERE Transaction_Type=1 AND(QestID NOT IN (SELECT QestID FROM  " & " InstallmentDet_Junc_Receipt WHERE Status <> 1))"
            StrSQL = StrSQL + ") as XTable"

            If Not IsNull(m_FromDate) Then
                StrSQL = StrSQL + " Where  DueDate >= #" & SQLDate(CDate(m_FromDate)) & "#"
                BolBegain = True
            End If

            If Not IsNull(m_ToDate) Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  DueDate <= #" & SQLDate(CDate(m_ToDate)) & "#"
                Else
                    StrSQL = StrSQL + " Where  DueDate <= #" & SQLDate(CDate(m_ToDate)) & "#"
                    BolBegain = True
                End If
            End If

            If LngCusID <> 0 Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  CusID =" & LngCusID & ""
                Else
                    StrSQL = StrSQL + " Where  CusID =" & LngCusID & ""
                    BolBegain = True
                End If
            End If

            StrSQL = StrSQL + " ORDER BY Transaction_ID,NoteID"
            Set RsCreditValues = New ADODB.Recordset
            RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        
            '«·√” ⁄·«„ ⁄‰ «·ﬁÌ„ «·„«·Ì… «·„” Õﬁ… ⁄·Ï «·‘—ﬂ… ··⁄„·«¡ Ê«·„Ê—œÌ‰
            StrSQL = "SELECT CompanyDebitValues.*"
            StrSQL = StrSQL + " FROM dbo.CompanyDebitValues() CompanyDebitValues "

            If Not IsNull(m_FromDate) Then
                StrSQL = StrSQL + " Where  DueDate >='" & SQLDate(CDate(m_FromDate)) & "'"
                BolBegain = True
            End If

            If Not IsNull(m_ToDate) Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  DueDate <='" & SQLDate(CDate(m_ToDate)) & "'"
                Else
                    StrSQL = StrSQL + " Where  DueDate <='" & SQLDate(CDate(m_ToDate)) & "'"
                    BolBegain = True
                End If
            End If

            If LngCusID <> 0 Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  CusID =" & LngCusID & ""
                Else
                    StrSQL = StrSQL + " Where  CusID =" & LngCusID & ""
                    BolBegain = True
                End If
            End If
        
            Set RsCreditValues = New ADODB.Recordset
            RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
    
    ElseIf IntType = 1 Then

        'Credit Values
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT Transaction_ID as TransactionsID, Transaction_Type, TransactionTypeName, Transaction_Date, CusName," & "NoteID, NoteType, Note_Value, DueDate, Transaction_Serial, CusID, NoteDate, NotesTypeName," & "PreRelease, Note_Value-IIF(IsNULL(PreRelease),0,PreRelease) AS RequiredValue " & " From  ( "
            StrSQL = StrSQL + " SELECT PaymentTime.Transaction_ID, PaymentTime.Transaction_Type," & "PaymentTime.TransactionTypeName, PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, " & "PaymentTime.NoteType, PaymentTime.Note_Value, PaymentTime.DueDate, PaymentTime.Transaction_Serial, " & "PaymentTime.CusID, PaymentTime.NoteDate, PaymentTime.NotesTypeName, Sum(QryPostedTransNotes.Note_Value) AS PreRelease" & " FROM PaymentTime LEFT JOIN QryPostedTransNotes ON PaymentTime.Transaction_ID =  " & " QryPostedTransNotes.Transaction_ID"
            StrSQL = StrSQL + " WHERE (((PaymentTime.Transaction_Type)=2 Or (PaymentTime.Transaction_Type)=5) " & "AND ((PaymentTime.NoteID) Not In (select NoteID From InstallMent) And (PaymentTime.NoteID)  "
            StrSQL = StrSQL + " Not In (select NoteID From TblCheckRelease)) "

            If StrPostedID <> "" Then
                StrSQL = StrSQL + " and ((PaymentTime.Transaction_ID) Not IN(" & StrPostedID & "))) "
            End If

            StrSQL = StrSQL + " GROUP BY PaymentTime.Transaction_ID, PaymentTime.Transaction_Type, PaymentTime.TransactionTypeName," & "PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, PaymentTime.NoteType, PaymentTime.Note_Value," & "PaymentTime.DueDate, PaymentTime.Transaction_Serial, PaymentTime.CusID, PaymentTime.NoteDate, " & "PaymentTime.NotesTypeName, PaymentTime.NotesTypeName "
    
            StrSQL = StrSQL + " Union "
    
            StrSQL = StrSQL + " SELECT   Transaction_ID,Transaction_Type,TransactionTypeName,Transaction_Date, "
            StrSQL = StrSQL + "CusName,NoteID,3 as NoteType,[Value] as Note_Value,DueDate,Transaction_Serial,CustID," & "Transaction_Date,'ﬁ”ÿ „” Õﬁ' as NotesTypeName,Summition as  PreRelease "
            StrSQL = StrSQL + " From QryCust_Qest"
            StrSQL = StrSQL + " WHERE Transaction_Type=2 AND(QestID NOT IN (SELECT QestID FROM  " & " InstallmentDet_Junc_Receipt WHERE Status <> 1))"
            StrSQL = StrSQL + ") as XTable"
        
            If Not IsNull(m_FromDate) Then
                StrSQL = StrSQL + " Where  DueDate >= #" & SQLDate(CDate(m_FromDate)) & "#"
                BolBegain = True
            End If

            If Not IsNull(m_ToDate) Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  DueDate <= #" & SQLDate(CDate(m_ToDate)) & "#"
                Else
                    StrSQL = StrSQL + " Where  DueDate <= #" & SQLDate(CDate(m_ToDate)) & "#"
                    BolBegain = True
                End If
            End If

            If LngCusID <> 0 Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  CusID =" & LngCusID & ""
                Else
                    StrSQL = StrSQL + " Where  CusID =" & LngCusID & ""
                    BolBegain = True
                End If
            End If

            StrSQL = StrSQL + " ORDER BY Transaction_ID,NoteID"
            Set RsCreditValues = New ADODB.Recordset
            RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            '«·√” ⁄·«„ ⁄‰ «·ﬁÌ„ ««·„«·Ì… «·„” Õﬁ… ··‘—ﬂ… ⁄·Ï «·⁄„·«¡ Ê«·„Ê—œÌ‰
            StrSQL = "SELECT CompanyCreditValues.* "
            StrSQL = StrSQL + " FROM dbo.CompanyCreditValues() CompanyCreditValues"

            If Not IsNull(m_FromDate) Then
                StrSQL = StrSQL + " Where  DueDate >='" & SQLDate(CDate(m_FromDate)) & "'"
                BolBegain = True
            End If

            If Not IsNull(m_ToDate) Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  DueDate <='" & SQLDate(CDate(m_ToDate)) & "'"
                Else
                    StrSQL = StrSQL + " Where  DueDate <='" & SQLDate(CDate(m_ToDate)) & "'"
                    BolBegain = True
                End If
            End If

            If LngCusID <> 0 Then
                If BolBegain = True Then
                    StrSQL = StrSQL + " AND  CusID =" & LngCusID & ""
                Else
                    StrSQL = StrSQL + " Where  CusID =" & LngCusID & ""
                    BolBegain = True
                End If
            End If
        
            Set RsCreditValues = New ADODB.Recordset
            RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If

    Set GetDebitCreditValues = RsCreditValues
End Function

Public Function UniqueNoteSerial1(NoteSerial1 As String, _
                                  IntTransType As Integer, _
                                  Optional LngTransID As Long = 0, _
                                  Optional intBranchId As Integer = 0) As Boolean

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT Transactions.Transaction_ID,Transactions.NoteSerial1, Transactions.Transaction_Type "
    StrSQL = StrSQL + " FROM Transactions "
    StrSQL = StrSQL + " Where NoteSerial1='" & NoteSerial1 & "'"
    StrSQL = StrSQL + " AND Transaction_Type=" & IntTransType & ""

    If LngTransID <> 0 Then
        StrSQL = StrSQL + " AND Transaction_ID <>" & LngTransID & ""
    End If

    If intBranchId <> 0 Then
        StrSQL = StrSQL + " AND BranchId =" & intBranchId & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.BOF) Then
        UniqueNoteSerial1 = False
    Else
        UniqueNoteSerial1 = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function UniqueTransSerial(StrTransSerial As String, _
                                  IntTransType As Integer, _
                                  Optional LngTransID As Long = 0, _
                                  Optional intBranchId As Integer = 0) As Boolean

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT Transactions.Transaction_ID,Transactions.Transaction_Serial, Transactions.Transaction_Type "
    StrSQL = StrSQL + " FROM Transactions "
    StrSQL = StrSQL + " Where Transaction_Serial='" & StrTransSerial & "'"
    StrSQL = StrSQL + " AND Transaction_Type=" & IntTransType & ""

    If LngTransID <> 0 Then
        StrSQL = StrSQL + " AND Transaction_ID <>" & LngTransID & ""
    End If

    If intBranchId <> 0 Then
        StrSQL = StrSQL + " AND BranchId =" & intBranchId & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.BOF) Then
        UniqueTransSerial = False
    Else
        UniqueTransSerial = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetCustomerAccount(LngCusID As Long, _
                                   Optional BolAllBalance As Boolean, _
                                   Optional m_ToDate As Variant, _
                                   Optional LngNotTransID As Long = 0) As Single

    Dim rs               As ADODB.Recordset
    Dim StrSQL           As String
    Dim SngBegainAccount As Single
    Dim SngCusAccount    As Single
    Dim BolBegin         As Boolean
    Dim RsNotes          As ADODB.Recordset
    Dim StrSQLNotes      As String
    Dim Account_code     As String
    '-----------------------------
    'Updated By ayman In 10-1-2008
    '-----------------------------
    StrSQL = "SELECT Account_Code,  TblCustemers.OpenBalance, TblCustemers.OpenBalanceType, " & "TblCustemers.CusID FROM TblCustemers "
    StrSQL = StrSQL + " Where  TblCustemers.CusID= " & LngCusID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        SngBegainAccount = IIf(IsNull(rs("OpenBalance").value), 0, rs("OpenBalance").value)
    End If

    If SngBegainAccount <> 0 Then
        If rs("OpenBalanceType").value = 0 Then
            SngBegainAccount = SngBegainAccount * -1
        End If
    End If

    rs.Close

    If BolAllBalance = True Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryCustomerBalance.CusID, QryCustomerBalance.CusName," & "SUM(QryCustomerBalance.Note_Value * QryCustomerBalance.CreditOrDebit)" & "as CustomerAccount"
            StrSQL = StrSQL + " From QryCustomerBalance "
            StrSQL = StrSQL + " Where QryCustomerBalance.CusID=" & LngCusID & ""

            If Not IsMissing(m_ToDate) Then
                StrSQL = StrSQL + " AND NoteDate <" & SQLDate(CDate(m_ToDate), True) & ""
            End If

            StrSQL = StrSQL + " Group BY QryCustomerBalance.CusID, QryCustomerBalance.CusName"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                SngCusAccount = IIf(IsNull(rs("CustomerAccount").value), 0, rs("CustomerAccount").value)
            End If

        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT QryCustomerBalance.CusID, QryCustomerBalance.CusName," & "SUM(QryCustomerBalance.Note_Value * QryCustomerBalance.CreditOrDebit)" & "as CustomerAccount"
            StrSQL = StrSQL + " From dbo.QryCustomerBalance(" & LngCusID & ")QryCustomerBalance "
            StrSQL = StrSQL + " Where QryCustomerBalance.CusID=" & LngCusID & ""

            If Not IsMissing(m_ToDate) Then
                StrSQL = StrSQL + " AND NoteDate <" & SQLDate(CDate(m_ToDate), True) & ""
            End If

            If LngNotTransID <> 0 Then
                Set RsNotes = New ADODB.Recordset
                StrSQLNotes = "Select NoteID From Notes Where Transaction_ID=" & LngNotTransID & ""
                StrSQLNotes = StrSQLNotes + "Order By NoteID"
                RsNotes.Open StrSQLNotes, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsNotes.BOF Or RsNotes.EOF) Then
                    RsNotes.MoveLast
                    StrSQL = StrSQL + " AND QryCustomerBalance.NoteID < " & RsNotes("NoteID").value & ""
                End If

                RsNotes.Close
                Set RsNotes = Nothing
            End If

            StrSQL = StrSQL + " Group BY QryCustomerBalance.CusID, QryCustomerBalance.CusName"
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                SngCusAccount = IIf(IsNull(rs("CustomerAccount").value), 0, rs("CustomerAccount").value)
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    GetCustomerAccount = SngBegainAccount + SngCusAccount

End Function

Public Function ShowRelatedNotes(LngTransID As Long, _
                                 IntType As Integer, _
                                 Optional ByRef SngValues As Single) As Boolean

    Dim StrSQL  As String
    Dim rs      As ADODB.Recordset
    Dim i       As Integer
    Dim Frm     As FrmShowTransNotes
    Dim RsTrans As ADODB.Recordset
    On Error GoTo hErr
    ShowRelatedNotes = False

    If LngTransID = 0 Then
        Exit Function
    End If

    'IntType =Request Only
    'IntType=Show Notes
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteSerial, dbo.Notes.NoteDate, dbo.Notes.NoteType," & "dbo.TblNotesTypes.NotesTypeName, dbo.Notes.Note_Value,dbo.Notes.CusID , dbo.TblCustemers.CusName," & "dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, dbo.Notes.Remark "
        StrSQL = StrSQL + " FROM dbo.Notes INNER JOIN dbo.TblNotesTypes ON dbo.Notes.NoteType =" & "dbo.TblNotesTypes.NotesType INNER JOIN dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID " & "LEFT OUTER JOIN dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID "
        StrSQL = StrSQL + " Where Transaction_ID=" & LngTransID
        StrSQL = StrSQL + " AND (dbo.Notes.NoteType=4 OR dbo.Notes.NoteType=9 Or " & "dbo.Notes.NoteType=5 OR dbo.Notes.NoteType=10)"
        StrSQL = StrSQL + " Order By dbo.Notes.NoteID"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Notes.NoteID, Notes.NoteSerial, Notes.NoteDate, Notes.NoteType," & "TblNotesTypes.NotesTypeName, Notes.Note_Value,Notes.CusID , TblCustemers.CusName,Notes.BoxID," & "TblBoxesData.BoxName, Notes.Remark "
        StrSQL = StrSQL + " FROM TblCustemers RIGHT JOIN (TblBoxesData RIGHT JOIN (TblNotesTypes INNER " & "JOIN Notes ON TblNotesTypes.NotesType = Notes.NoteType) ON TblBoxesData.BoxID = Notes.BoxID) ON " & "TblCustemers.CusID = Notes.CusID "
        StrSQL = StrSQL + " Where Transaction_ID=" & LngTransID
        StrSQL = StrSQL + " And (Notes.NoteType = 4 Or Notes.NoteType = 9 Or Notes.NoteType = 5 Or " & "Notes.NoteType = 10)"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        ShowRelatedNotes = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    If IntType = 0 Then
        If Not IsMissing(SngValues) Then

            For i = 0 To rs.RecordCount - 1
                SngValues = SngValues + IIf(IsNull(rs("Note_Value").value), 0, rs("Note_Value").value)
                rs.MoveNext
            Next i

        End If

        rs.Close
        Set rs = Nothing
        ShowRelatedNotes = True
        Exit Function
    End If

    '--------------------------------------------------------------------------------------------------
    Set Frm = New FrmShowTransNotes
    Frm.IntMode = 0

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusID,dbo.TblCustemers.CusName ," & "dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName "
        StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN  dbo.TblCustemers ON " & "dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN dbo.TransactionTypes ON " & "dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
        StrSQL = StrSQL + " Where  dbo.Transactions.Transaction_ID=" & LngTransID & ""
    
        Debug.Print Replace(StrSQL, "dbo.", "", , , vbTextCompare)
    
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT     Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Date, TblCustemers.CusID,TblCustemers.CusName," & "Transactions.Transaction_Type, TransactionTypes.TransactionTypeName  "
        StrSQL = StrSQL + " FROM (Transactions INNER JOIN  TblCustemers ON Transactions.CusID =" & "TblCustemers.CusID) INNER JOIN TransactionTypes ON Transactions.Transaction_Type = " & "TransactionTypes.Transaction_Type "
        StrSQL = StrSQL + " Where  Transactions.Transaction_ID=" & LngTransID & ""
    End If

    Set RsTrans = New ADODB.Recordset
    RsTrans.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTrans.BOF Or RsTrans.EOF) Then
        Frm.lbl(7).Caption = IIf(IsNull(RsTrans("Transaction_Serial").value), "", RsTrans("Transaction_Serial").value)
        Frm.lbl(9).Caption = IIf(IsNull(RsTrans("TransactionTypeName").value), "", RsTrans("TransactionTypeName").value)

        If Not IsNull(RsTrans("Transaction_Date").value) Then
            Frm.lbl(12).Caption = DisplayDate(RsTrans("Transaction_Date").value)
        End If

        Frm.LblLink.Caption = IIf(IsNull(RsTrans("CusName").value), "", RsTrans("CusName").value)
        Frm.LblLink.Tag = IIf(IsNull(RsTrans("CusID").value), "", RsTrans("CusID").value)
    End If

    RsTrans.Close
    Set RsTrans = Nothing

    '--------------------------------------------------------------------------------------------------
    With Frm.fg
        .rows = .FixedRows + rs.RecordCount

        For i = 1 To .rows - 1
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

            If Not (IsNull(rs("NoteDate").value)) Then
                .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(rs("NoteDate").value)
            End If

            .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(rs("NotesTypeName").value), "", rs("NotesTypeName").value)
            .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(rs("NoteType").value), "", rs("NoteType").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(rs("Note_Value").value), 0, rs("Note_Value").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
            .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
            rs.MoveNext
        Next i

        Frm.lbl(5).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("NoteID"), .rows, .ColIndex("NoteID"))
        Frm.lbl(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Note_Value"), .rows, .ColIndex("Note_Value"))
        .AutoSize 0, .Cols - 1, False
    End With

    Frm.show
    Exit Function
hErr:
    ShowRelatedNotes = False
End Function

Public Function CheckCusCredit(LngCusID As Long, _
                               SngOutValue As Single, _
                               IntCheckType As Integer, _
                               Optional Transaction_ID As Double) As Boolean

    Dim rs                   As ADODB.Recordset
    Dim StrSQL               As String
    Dim SngCreditLiimt       As Single
    Dim SngCreditLimitCredit As Single
    Dim SngCusAccount        As Single
    Dim Msg                  As String
    Dim StrTemp              As String
    Dim IntRes               As Integer
    On Local Error GoTo ErrTrap

    'IntCheckType= 0 Check For Debit
    'IntCheckType= 1 Check For Credit

    StrSQL = "Select Account_Code,CreditLimit,CreditLimitCredit From TblCustemers Where CusID=" & LngCusID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        SngCreditLiimt = IIf(IsNull(rs("CreditLimit").value), 0, rs("CreditLimit").value)
        SngCreditLimitCredit = IIf(IsNull(rs("CreditLimitCredit").value), 0, rs("CreditLimitCredit").value)
    Else
        CheckCusCredit = False
        Exit Function
    End If

    If IntCheckType = 0 Then

        '«·ﬂ‘› ⁄·Ï «‰ „œÌÊ‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
        If SngCreditLiimt = 0 Then
            'NO CreditLimit For this customer
            CheckCusCredit = True
            Exit Function
        Else
            'Set Rs = New ADODB.Recordset
            '            SngCusAccount = GetCustomerAccount(LngCusID, True)
 
            '------------------------------------------------
            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
            '*******new code********************************************
            Dim Account_code As String
            Dim FirstPeriod  As Date
            getFirstPeriodDateInthisYear FirstPeriod
        
            Account_code = GetMyAccountCode("TblCustemers", "CusID", LngCusID)  '
            SngCusAccount = GetActualAccountBalance(Account_code, 0, FirstPeriod, Date)
 
            SngCusAccount = SngCusAccount - GetSumOfGeForOneAccount(Account_code, Transaction_ID, 0)
  
            '***************************************************\
            
            If SngCusAccount >= 0 Then '„œÌ‰
                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLiimt Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                        Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ «·Œ«’ »«·⁄„Ì·...!!!"
                        Msg = Msg & CHR(13) & "------------------------------------------------"
                        Msg = Msg & CHR(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
                        Msg = Msg & CHR(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì·  ﬁ»· Â–Â «·Õ—ﬂ…: "
                    Else
             
                        Msg = "Can't Allow to complete this Transaction"
                        Msg = Msg & CHR(13) & " Over Limit !!!"
                        Msg = Msg & CHR(13) & "------------------------------------------------"
                        Msg = Msg & CHR(13) & "Limit is :  : " & SngCreditLiimt & " " & WriteNo(CStr(SngCreditLiimt), 0)
                        Msg = Msg & CHR(13) & " Current Balance Is: "
                    
                    End If
                    If SngCusAccount > 0 Then
                        StrTemp = Abs(SngCusAccount) & "(„œÌ‰)"
                    ElseIf SngCusAccount < 0 Then
                        StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
                    Else
                        StrTemp = "(Œ«·’)"
                    End If

                    Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
                    Msg = Msg & CHR(13) & ""
                 
                    If SystemOptions.AllowCreditPass = False Then
                        Msg = Msg & CHR(13) & " ·« Ì„ﬂ‰ «·≈” „—«— ›Ï Õ›Ÿ «·›« Ê—… ...øøø"

                        IntRes = MsgBox(Msg, vbExclamation + vbNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
                        CheckCusCredit = False
                        Exit Function
                    Else
                        Msg = Msg & CHR(13) & " Â·  —Ìœ «·≈” „—«— ›Ï Õ›Ÿ «·›« Ê—… ...øøø"

                        IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
                    End If
                    
                    If IntRes = vbNo Then
                        CheckCusCredit = False
                        Exit Function
                    Else
                        CheckCusCredit = True
                        Exit Function
                    End If
                    
                End If
            End If

            '------------------------------------------------
        End If

    ElseIf IntCheckType = 1 Then

        '«·ﬂ‘› ⁄·Ï «‰ œ«∆‰Ì… «·⁄„Ì· ·‰  “Ìœ ⁄‰ «·Õœ «·„Õœœ ·Â
        If SngCreditLimitCredit = 0 Then
            'NO CreditLimit For this customer
            CheckCusCredit = True
            Exit Function
        Else
            'Set Rs = New ADODB.Recordset
            SngCusAccount = GetCustomerAccount(LngCusID, True)

            '------------------------------------------------
            '»⁄œ «·√” ⁄·«„ ⁄‰ —’Ìœ «·⁄„Ì·
            If SngCusAccount >= 0 Then '„œÌ‰
                If (Abs(SngCusAccount) + SngOutValue) > SngCreditLimitCredit Then
                    Msg = "·«Ì„ﬂ‰ «·”„«Õ »Â–Â «·⁄„·Ì… ...!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê›   ŒÿÏ Õœ «·√∆ „«‰ («·œ«∆‰) «·Œ«’ »«·⁄„Ì· ...!!!"
                    Msg = Msg & CHR(13) & "------------------------------------------------"
                    Msg = Msg & CHR(13) & "Õœ ≈∆ „«‰ «·⁄„Ì· : " & SngCreditLimitCredit & " " & WriteNo(CStr(SngCreditLimitCredit), 0)
                    Msg = Msg & CHR(13) & "«·—’Ìœ «·Õ«·Ï ··⁄„Ì· : "

                    If SngCusAccount < 0 Then
                        StrTemp = Abs(SngCusAccount) & "(œ«∆‰)"
                    Else
                        StrTemp = "(Œ«·’)"
                    End If

                    Msg = Msg & StrTemp
                    Msg = Msg & CHR(13) & "«·„»·€ «·„—«œ  ”ÃÌ·Â ⁄·Ï «·⁄„Ì· : " & SngOutValue
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CheckCusCredit = False
                    Exit Function
                End If
            End If

            '------------------------------------------------
        End If
    End If

    CheckCusCredit = True
    Exit Function
ErrTrap:
    CheckCusCredit = False
End Function
Public Function GetItemQuantityStock(LngItemID As Long, _
                                     Optional LngStoreID As Long = 0, _
                                     Optional DtDate, _
                                     Optional LngTransID As Long = 0, _
                                     Optional BolHasExpire As Boolean = False, _
                                     Optional StrItem_Name As String = "", _
                                     Optional StrItemSerial As String = "", _
                                     Optional AllSerails As Boolean = False, _
                                     Optional LngColorID As Long = 1, _
                                     Optional StrItemSize As String = "", _
                                     Optional ClassId As Long = 1) As ADODB.Recordset
            
    'Â–« «·œ«·… Â«„… Ãœ« ›Ï «·»—‰«„Ã
    'ÕÌÀ «‰Â«  ﬁÊ„ »≈⁄ÿ«¡ ﬂ„Ì… «·’‰›
    '«Ê «·√’‰«› «·„ÊÃÊœ… ›Ï «·»—‰«„Ã
    '»„⁄‰Ï «‰Â« «· Ï  ﬁÊ„ »⁄„· ⁄„·Ì…
    '«·Ã—œ ›Ï «·»—‰«„Ã

    Dim RsSelect As ADODB.Recordset
    Dim XCmd     As ADODB.Command
    Dim XPar     As ADODB.Parameter
    Dim YPar     As ADODB.Parameter
    Dim StrSQL   As String
    Dim DatePar  As Date

    On Error GoTo ErrTrap

    If LngItemID = 0 Then
        Set GetItemQuantityStock = Nothing
        Exit Function
    End If

    If IsMissing(DtDate) Then
        DatePar = Date
    Else
        DatePar = DtDate
    End If

    If StrItemSerial = "" Then

        'Check For all Quantity
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "Select * From QryQuantity Where ItemID=" & LngItemID & ""

            If LngStoreID <> 0 Then
                StrSQL = StrSQL & " And StoreID=" & LngStoreID & ""
            End If

            Set XCmd = New ADODB.Command
            Set XCmd.ActiveConnection = Cn
            Set XPar = New ADODB.Parameter

            With XPar
                .Name = "ToDateX"
                .Direction = adParamInput
                .type = adDate
                '.Value = SQLDate(DatePar)
                .value = DatePar
                XCmd.Parameters.Append XPar
            End With

            Set YPar = New ADODB.Parameter

            With YPar
                .Name = "LngTransID"
                .Direction = adParamInput
                .type = adInteger
                .value = LngTransID
                XCmd.Parameters.Append YPar
            End With

            XCmd.CommandType = adCmdText
            XCmd.CommandText = StrSQL
            Set RsSelect = New ADODB.Recordset
            RsSelect.CursorLocation = adUseClient
            RsSelect.CursorType = adOpenStatic
            RsSelect.LockType = adLockReadOnly
            RsSelect.Open XCmd, , adOpenStatic, adLockReadOnly
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
   
            StrSQL = "SELECT  sum(qty) As totalqty " & "FROM dbo.QryQuantity('" & SQLDate(DatePar, False) & "'," & LngTransID & ") QryQuantity"
        
            '     StrSQL = "SELECT QryQuantity.* " & _
                  " FROM dbo.QryQuantity('" & SQLDate(DatePar, False) & "'," & LngTransID & ") QryQuantity"
            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""

            If LngStoreID <> 0 Then
                StrSQL = StrSQL + " AND QryQuantity.StoreID=" & LngStoreID & ""
            End If

            StrSQL = StrSQL + " and ColorID=" & LngColorID
            'If StrItemSize = "" Then StrSQL = StrSQL + " and itemSize Is Null" Else StrSQL = StrSQL + " and '" & StrItemSize & "'"

            StrSQL = StrSQL + " and ItemSize=" & StrItemSize
        
            'StrSQL = StrSQL + " and (ItemSize='" & ItemSize & "' Or ItemSize Is Null)"
            
            'cancelled
           
            StrSQL = " SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS totalqty"
            StrSQL = StrSQL + " FROM         dbo.Transaction_Details INNER JOIN"
            StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
            StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
            StrSQL = StrSQL + " WHERE   ( 1 = 1 "
            If LngStoreID <> 0 Then
                StrSQL = StrSQL + "  and  dbo.Transactions.StoreID  =" & LngStoreID
            End If
            If LngTransID <> 0 Then
                StrSQL = StrSQL + " and dbo.Transactions.Transaction_ID <> " & LngTransID
            End If
            StrSQL = StrSQL + " AND  dbo.Transactions.Transaction_Date <=' " & SQLDate(DatePar, False)
            StrSQL = StrSQL + "'  AND (dbo.TransactionTypes.StockEffect <> 0)   "
            If LngColorID <> 0 Then
            StrSQL = StrSQL + "AND (dbo.Transaction_Details.ColorID = " & LngColorID & ") "
            End If
            If StrItemSize <> "" Then
                StrSQL = StrSQL + " AND (dbo.Transaction_Details.ItemSize = N'" & StrItemSize & "') "
            End If
            If ClassId <> 1 Then
            StrSQL = StrSQL + " AND (dbo.Transaction_Details.ClassId = " & ClassId & ")"
            End If
            StrSQL = StrSQL + "  ) GROUP BY dbo.Transaction_Details.Item_ID"
            StrSQL = StrSQL + "  Having (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
            
            Set RsSelect = New ADODB.Recordset
            RsSelect.CursorLocation = adUseClient
            RsSelect.CursorType = adOpenStatic
            RsSelect.LockType = adLockReadOnly
            RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If

    ElseIf StrItemSerial <> "" Then

        If AllSerails = False Then '≈–« ﬂ‰   —Ìœ «·√” ⁄·«„ ⁄‰ ”Ì—Ì«· „Õœœ
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = "Select * From QryGardComplete"
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & " AND ItemSerial='" & StrItemSerial & "'"
                StrSQL = StrSQL + " and StoreID=" & LngStoreID
                Set RsSelect = New ADODB.Recordset
                RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                '  StrSQL = "Select QryQuantity.* From dbo.QryQuantity('" & SQLDate(DatePar, False) & "'," & LngTransID & ")QryQuantity "
                '  StrSQL = StrSQL + " Where ItemID=" & LngItemID & _
                '  " AND ItemSerial='" & StrItemSerial & "'"
                '  StrSQL = StrSQL + " and StoreID=" & LngStoreID
                '  StrSQL = StrSQL + " and ColorID=" & LngColorID
                '  StrSQL = StrSQL + " and ItemSize='" & StrItemSize & "'"
            
                StrSQL = " SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS totalqty"
                StrSQL = StrSQL + " FROM         dbo.Transaction_Details INNER JOIN"
                StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
                StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
                StrSQL = StrSQL + " WHERE     ( dbo.Transactions.StoreID  =" & LngStoreID & " and dbo.Transactions.Transaction_ID <>" & LngTransID & " AND  dbo.Transactions.Transaction_Date <=' " & SQLDate(DatePar, False) & "'  AND (dbo.TransactionTypes.StockEffect <> 0) AND"
                StrSQL = StrSQL + " (dbo.Transaction_Details.ColorID = " & LngColorID & ") "
                If LngTransID <> 0 Then
                    StrSQL = StrSQL & " and dbo.Transactions.Transaction_ID <>" & LngTransID
                End If
                If StrItemSize <> "" Then
                    StrSQL = StrSQL + " AND (dbo.Transaction_Details.ItemSize = N'" & StrItemSize & "') "
                End If
                
                StrSQL = StrSQL + " AND (dbo.Transaction_Details.ClassId = " & ClassId & ")"
                If StrItemSerial <> "" Then
                    StrSQL = StrSQL + "  AND (dbo.Transaction_Details.ItemSerial = N'" & StrItemSerial & "')"
                End If
                StrSQL = StrSQL + "  ) GROUP BY dbo.Transaction_Details.Item_ID"
                StrSQL = StrSQL + "  Having (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
            
                Set RsSelect = New ADODB.Recordset
                RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            End If

        ElseIf AllSerails = True Then

            '≈–« ﬂ‰   —Ìœ ﬂ· «·”Ì—Ì«·«  «·„ÊÃÊœ…
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = "Select * From QryGardComplete"
                StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                StrSQL = StrSQL + " and StoreID=" & LngStoreID
                Set RsSelect = New ADODB.Recordset
                RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                '            StrSQL = "Select QryQuantity.* From dbo.QryQuantity('" & SQLDate(DatePar, False) & "'," & LngTransID & ")QryQuantity "
                '            StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
                '            StrSQL = StrSQL + " And StoreID=" & LngStoreID
                '            StrSQL = StrSQL + " and ColorID=" & LngColorID
                '            StrSQL = StrSQL + " and ItemSize='" & StrItemSize & "'"
                StrSQL = " SELECT     SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS totalqty"
                StrSQL = StrSQL + " FROM         dbo.Transaction_Details INNER JOIN"
                StrSQL = StrSQL + " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
                StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
                StrSQL = StrSQL + " WHERE     ( dbo.Transactions.StoreID  =" & LngStoreID & " and dbo.Transactions.Transaction_ID <>" & LngTransID & " AND  dbo.Transactions.Transaction_Date <=' " & SQLDate(DatePar, False) & "'  AND (dbo.TransactionTypes.StockEffect <> 0) AND"
                StrSQL = StrSQL + " (dbo.Transaction_Details.ColorID = " & LngColorID & ") "
                If StrItemSize <> "" Then
                    StrSQL = StrSQL + " AND (dbo.Transaction_Details.ItemSize = N'" & StrItemSize & "') "
                End If
               
                StrSQL = StrSQL + " AND (dbo.Transaction_Details.ClassId = " & ClassId & ")"
                StrSQL = StrSQL + "  AND (dbo.Transaction_Details.ItemSerial = N'" & StrItemSerial & "')"
                StrSQL = StrSQL + "  ) GROUP BY dbo.Transaction_Details.Item_ID"
                StrSQL = StrSQL + "  Having (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
            
                Set RsSelect = New ADODB.Recordset
                RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            End If
        End If
    End If

    Set GetItemQuantityStock = RsSelect
Exit_Function:
    Exit Function
ErrTrap:
    Set GetItemQuantityStock = Nothing
    'If AllSerails = False Then '≈–« ﬂ‰   —Ìœ «·√” ⁄·«„ ⁄‰ ”Ì—Ì«· „Õœœ
    '    If SystemOptions.SysDataBaseType = AccessDataBase Then
    '        StrSQL = "Select * From QryGardComplete"
    '        StrSQL = StrSQL + " Where ItemID=" & LngItemID & _
    '        " AND ItemSerial='" & StrItemSerial & "'"
    '        StrSQL = StrSQL + " and StoreID=" & LngStoreID
    '        Set RsSelect = New ADODB.Recordset
    '        RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '        StrSQL = "Select * From dbo.QryGardComplete(" & LngTransID & ")QryGardComplete "
    '        StrSQL = StrSQL + " Where ItemID=" & LngItemID & _
    '        " AND ItemSerial='" & StrItemSerial & "'"
    '        StrSQL = StrSQL + " and StoreID=" & LngStoreID
    '        Set RsSelect = New ADODB.Recordset
    '        RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    End If
    '    ElseIf AllSerails = True Then
    '    '≈–« ﬂ‰   —Ìœ ﬂ· «·”Ì—Ì«·«  «·„ÊÃÊœ…
    '    If SystemOptions.SysDataBaseType = AccessDataBase Then
    '        StrSQL = "Select * From QryGardComplete"
    '        StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
    '        StrSQL = StrSQL + " and StoreID=" & LngStoreID
    '        Set RsSelect = New ADODB.Recordset
    '        RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '        StrSQL = "Select * From dbo.QryGardComplete(" & LngTransID & ")QryGardComplete "
    '        StrSQL = StrSQL + " Where ItemID=" & LngItemID & ""
    '        StrSQL = StrSQL + " And StoreID=" & LngStoreID
    '        Set RsSelect = New ADODB.Recordset
    '        RsSelect.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '    End If
End Function
 
Public Sub ReCusDCombo(MyCombo As DataCombo, _
   IntType As Integer)
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers IntType, MyCombo, True
    Set Dcombos = Nothing
End Sub

Public Sub ShowCusBalDailog(LngCusID As Long, _
   IntReportIndex As Integer)
    Load FrmSelectDate
    FrmSelectDate.Retrive LngCusID
    FrmSelectDate.CboReportType.ListIndex = IntReportIndex
    FrmSelectDate.show
    FrmSelectDate.ZOrder 0
End Sub

Public Function ShowCurrencyAlarm(Optional ShowMsg As Boolean = False) As Boolean
    On Error GoTo ErrTrap
    Dim Msg            As String
    Dim StrSQL         As String
    Dim RsTemp         As New ADODB.Recordset
    Dim RsTest         As New ADODB.Recordset
    Dim rs             As New ADODB.Recordset
    Dim RsCreditValues As ADODB.Recordset
    Dim RsDebitValues  As ADODB.Recordset
    Dim StrPostID      As String
    Dim StrCaption     As String
    Dim StrPostedID    As String
    Dim cProgress      As ClsProgress
    Dim i              As Integer
    Dim RsChartCredit  As ADODB.Recordset
    Dim RsChartDebit   As ADODB.Recordset
    Dim maskString     As String
    '-----------
    maskString = "%l " + vbCrLf + "%v" + "(" & "%p" + ")"
    '-----------
    StrPostedID = GetPostTransID

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        'Credit Values
        StrSQL = "SELECT Transaction_ID as TransactionsID, Transaction_Type, TransactionTypeName, Transaction_Date, CusName," & "NoteID, NoteType, Note_Value, DueDate, Transaction_Serial, CusID, NoteDate, NotesTypeName," & "PreRelease, Note_Value-IIF(IsNULL(PreRelease),0,PreRelease) AS RequiredValue " & " From  ( "
        StrSQL = StrSQL + " SELECT PaymentTime.Transaction_ID, PaymentTime.Transaction_Type," & "PaymentTime.TransactionTypeName, PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, " & "PaymentTime.NoteType, PaymentTime.Note_Value, PaymentTime.DueDate, PaymentTime.Transaction_Serial, " & "PaymentTime.CusID, PaymentTime.NoteDate, PaymentTime.NotesTypeName, Sum(QryPostedTransNotes.Note_Value) AS PreRelease" & " FROM PaymentTime LEFT JOIN QryPostedTransNotes ON PaymentTime.Transaction_ID =  " & " QryPostedTransNotes.Transaction_ID"
        StrSQL = StrSQL + " WHERE (((PaymentTime.Transaction_Type)=2 Or (PaymentTime.Transaction_Type)=5) " & "AND ((PaymentTime.NoteID) Not In (select NoteID From InstallMent) And (PaymentTime.NoteID)  "
        StrSQL = StrSQL + " Not In (select NoteID From TblCheckRelease)) "

        If StrPostedID <> "" Then
            StrSQL = StrSQL + " and ((PaymentTime.Transaction_ID) Not IN(" & StrPostedID & "))"
        End If

        StrSQL = StrSQL + ")"
        StrSQL = StrSQL + " GROUP BY PaymentTime.Transaction_ID, PaymentTime.Transaction_Type, PaymentTime.TransactionTypeName," & "PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, PaymentTime.NoteType, PaymentTime.Note_Value," & "PaymentTime.DueDate, PaymentTime.Transaction_Serial, PaymentTime.CusID, PaymentTime.NoteDate, " & "PaymentTime.NotesTypeName, PaymentTime.NotesTypeName "

        StrSQL = StrSQL + " Union "

        StrSQL = StrSQL + " SELECT   Transaction_ID,Transaction_Type,TransactionTypeName,Transaction_Date, "
        StrSQL = StrSQL + "CusName,NoteID,3 as NoteType,[Value] as Note_Value,DueDate,Transaction_Serial,CustID," & "Transaction_Date,'ﬁ”ÿ „” Õﬁ' as NotesTypeName,Summition as  PreRelease "
        StrSQL = StrSQL + " From QryCust_Qest"
        StrSQL = StrSQL + " WHERE Transaction_Type=2 AND(QestID NOT IN (SELECT QestID FROM  " & " InstallmentDet_Junc_Receipt WHERE Status <> 1))"
        StrSQL = StrSQL + ") as XTable"
        StrSQL = StrSQL + " Where  DueDate <= #" & SQLDate(Date) & "#"
        StrSQL = StrSQL + " ORDER BY NoteID"
        Set RsCreditValues = New ADODB.Recordset
        RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        '=========================================================
        'Debit Values
        StrSQL = "SELECT Transaction_ID as TransactionsID, Transaction_Type, TransactionTypeName, Transaction_Date, CusName," & "NoteID, NoteType, Note_Value, DueDate, Transaction_Serial, CusID, NoteDate, NotesTypeName," & "PreRelease, Note_Value-IIF(IsNULL(PreRelease),0,PreRelease) AS RequiredValue " & " From  ( "
        StrSQL = StrSQL + " SELECT PaymentTime.Transaction_ID, PaymentTime.Transaction_Type," & "PaymentTime.TransactionTypeName, PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, " & "PaymentTime.NoteType, PaymentTime.Note_Value, PaymentTime.DueDate, PaymentTime.Transaction_Serial, " & "PaymentTime.CusID, PaymentTime.NoteDate, PaymentTime.NotesTypeName, Sum(QryPostedTransNotes.Note_Value) AS PreRelease" & " FROM PaymentTime LEFT JOIN QryPostedTransNotes ON PaymentTime.Transaction_ID =  " & " QryPostedTransNotes.Transaction_ID"
        StrSQL = StrSQL + " WHERE (((PaymentTime.Transaction_Type)=1 Or (PaymentTime.Transaction_Type)=9) " & "AND ((PaymentTime.NoteID) Not In (select NoteID From InstallMent) And (PaymentTime.NoteID)  "
        StrSQL = StrSQL + " Not In (select NoteID From TblCheckRelease)) "

        If StrPostedID <> "" Then
            'AND ((PaymentTime.Transaction_ID) Not In (1,0)))"
            StrSQL = StrSQL + " and ((PaymentTime.Transaction_ID) Not IN(" & StrPostedID & ")) "
        End If

        StrSQL = StrSQL + ")"
        StrSQL = StrSQL + " GROUP BY PaymentTime.Transaction_ID, PaymentTime.Transaction_Type, PaymentTime.TransactionTypeName," & "PaymentTime.Transaction_Date, PaymentTime.CusName, PaymentTime.NoteID, PaymentTime.NoteType, PaymentTime.Note_Value," & "PaymentTime.DueDate, PaymentTime.Transaction_Serial, PaymentTime.CusID, PaymentTime.NoteDate, " & "PaymentTime.NotesTypeName, PaymentTime.NotesTypeName "

        StrSQL = StrSQL + " Union "

        StrSQL = StrSQL + " SELECT   Transaction_ID,Transaction_Type,TransactionTypeName,Transaction_Date, "
        StrSQL = StrSQL + "CusName,NoteID,3 as NoteType,[Value] as Note_Value,DueDate,Transaction_Serial,CustID," & "Transaction_Date,'ﬁ”ÿ „” Õﬁ' as NotesTypeName,Summition as  PreRelease "
        StrSQL = StrSQL + " From QryCust_Qest"
        StrSQL = StrSQL + " WHERE Transaction_Type=1 AND(QestID NOT IN (SELECT QestID FROM  " & " InstallmentDet_Junc_Receipt WHERE Status <> 1))"
        StrSQL = StrSQL + ") as XTable"
        StrSQL = StrSQL + " Where  DueDate <= #" & SQLDate(Date) & "#"
        StrSQL = StrSQL + " ORDER BY NoteID"
        Set RsDebitValues = New ADODB.Recordset
        RsDebitValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        '«·√” ⁄·«„ ⁄‰ «·ﬁÌ„ ««·„«·Ì… «·„” Õﬁ… ··‘—ﬂ… ⁄·Ï «·⁄„·«¡ Ê«·„Ê—œÌ‰

        Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_ShowPayment", True)
        Askcount = GetSetting(StrAppRegPath, "Setting", "count_ShowPayment", 0)

        StrSQL = "SELECT CompanyCreditValues.* "
        StrSQL = StrSQL + " FROM dbo.CompanyCreditValues() CompanyCreditValues"
        StrSQL = StrSQL + " Where  requiredvalue>0 and DueDate <= '" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        Set RsCreditValues = New ADODB.Recordset
        RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute + adAsyncFetch
        Debug.Print StrSQL
        '-----------------›Ï ÌÊ„ 17 ”» „»— 2007
        StrSQL = "Select Sum(RequiredValue)as XXX,COl1 " & "From(SELECT CompanyCreditValues.* ,DateDiff(day,DueDate," & SQLDate(Date, True) & ")as PastAgo " & ",'Col1'=Case    " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <= 10 THEN 10   " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <= 30 THEN 30   " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <" & "= 60 THEN 60  WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & " ) <= 90 THEN 90     ELSE        '1" & "0000'     END     FROM dbo.CompanyCreditValues() CompanyCreditValues      Where  DueDate " & "< " & SQLDate(Date, True) & " ) xTable Group By COL1 "
    
        Set RsChartCredit = New ADODB.Recordset
        RsChartCredit.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute + adAsyncFetch
        '------------------------------------
        '«·√” ⁄·«„ ⁄‰ «·ﬁÌ„ «·„«·Ì… «·„” Õﬁ… ⁄·Ï «·‘—ﬂ… ··⁄„·«¡ Ê«·„Ê—œÌ‰
        StrSQL = "SELECT CompanyDebitValues.*"
        StrSQL = StrSQL + " FROM dbo.CompanyDebitValues() CompanyDebitValues "
        StrSQL = StrSQL + " Where  requiredvalue>0 and DueDate <= '" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        Set RsDebitValues = New ADODB.Recordset
        RsDebitValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute + adAsyncFetch
        '-----------------›Ï ÌÊ„ 17 ”» „»— 2007
        StrSQL = "Select Sum(RequiredValue)as XXX,COl1 " & "From(SELECT CompanyDebitValues.* ,DateDiff(day,DueDate," & SQLDate(Date, True) & ")as PastAgo " & ",'Col1'=Case    " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <= 10 THEN 10   " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <= 30 THEN 30   " & "WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & ") <" & "= 60 THEN 60  WHEN DateDiff(day,DueDate," & SQLDate(Date, True) & " ) <= 90 THEN 90     ELSE        '1" & "0000'     END     FROM dbo.CompanyDebitValues() CompanyDebitValues      Where  DueDate " & "< " & SQLDate(Date, True) & " ) xTable Group By COL1 "
    
        Set RsChartDebit = New ADODB.Recordset
        RsChartDebit.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute + adAsyncFetch
        '------------------------------------
        Set cProgress = New ClsProgress
        cProgress.ProgressType = Waiting
    
        cProgress.StartProgress

        Do While RsCreditValues.State = adStateExecuting Or RsDebitValues.State = adStateExecuting Or RsChartCredit.State = adStateExecuting Or RsChartDebit.State = adStateExecuting

            DoEvents
        Loop

    End If

    If (RsDebitValues.BOF Or RsDebitValues.EOF) And (RsCreditValues.BOF Or RsCreditValues.EOF) Then
        'NO Debit Or Credit Values
        'The Recordsets are emptyes
        cProgress.StopProgess
        Set cProgress = Nothing

        If ShowMsg = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«  ÊÃœ √Ì „»«·€ Õ«‰ Êﬁ  «” ·«„Â« √Ê ”œ«œÂ«"
            Else
                Msg = " No Data "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If

        ShowCurrencyAlarm = False
        Exit Function
    Else
        Load FrmPaymentTime

        If Not (RsCreditValues.BOF Or RsCreditValues.EOF) Then

            With FrmPaymentTime.FG1
                .rows = .FixedRows + RsCreditValues.RecordCount
                RsCreditValues.MoveFirst

                For i = 1 To RsCreditValues.RecordCount
                    .TextMatrix(i, .ColIndex("TransactionsID")) = IIf(IsNull(RsCreditValues("TransactionsID").value), "", RsCreditValues("TransactionsID").value)
                    .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(RsCreditValues("Transaction_Type").value), "", RsCreditValues("Transaction_Type").value)
                    .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsCreditValues("TransactionTypeName").value), "", RsCreditValues("TransactionTypeName").value)
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsCreditValues("CusName").value), "", RsCreditValues("CusName").value)
                    .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsCreditValues("NoteID").value), "", RsCreditValues("NoteID").value)
                    .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(RsCreditValues("NoteType").value), "", RsCreditValues("NoteType").value)
                    .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsCreditValues("Note_Value").value), "", RsCreditValues("Note_Value").value)

                    If Not IsNull(RsCreditValues("DueDate").value) Then
                        .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsCreditValues("DueDate").value)
                        .TextMatrix(i, .ColIndex("LateInterval")) = DateDiff("d", RsCreditValues("DueDate").value, Date)
                    End If

                    .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsCreditValues("Transaction_Serial").value), "", RsCreditValues("Transaction_Serial").value)

                    If Not IsNull(RsCreditValues("Transaction_Date").value) Then
                        .TextMatrix(i, .ColIndex("Transaction_Date")) = DisplayDate(RsCreditValues("Transaction_Date").value)
                    End If

                    .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsCreditValues("CusID").value), "", RsCreditValues("CusID").value)

                    If Not IsNull(RsCreditValues("NoteDate").value) Then
                        .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(RsCreditValues("NoteDate").value)
                    End If

                    .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(RsCreditValues("NotesTypeName").value), "", RsCreditValues("NotesTypeName").value)
                    .TextMatrix(i, .ColIndex("PreRelease")) = IIf(IsNull(RsCreditValues("PreRelease").value), "", RsCreditValues("PreRelease").value)
                    .TextMatrix(i, .ColIndex("RequiredValue")) = IIf(IsNull(RsCreditValues("RequiredValue").value), "", RsCreditValues("RequiredValue").value)
                    RsCreditValues.MoveNext
                Next i

                DrawFloodProgress FrmPaymentTime.FG1, .ColIndex("LateInterval")
                .AutoSize 0, .Cols - 1, False
            End With

        End If

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            If Not (RsChartCredit.BOF Or RsChartCredit.EOF) Then
                '            With FrmPaymentTime.ChartRecv
                '                .PointLabelMask = maskString
                '                .PointLabels = False
                '                .Titles(0).text = "„ «Œ—«  «· Õ’Ì· „‰ «·⁄„·«¡"
                '                .Titles(0).Alignment = StringAlignment_Near
                '                .Titles(0).LineAlignment = StringAlignment_Near
                '                .Titles(0).Font.Bold = True
                '                .Titles(0).Font.Charset = 178
                '                .Titles(0).DockArea = DockArea_Top
                '                .Titles(0).DrawingArea = True
                '                .Titles(0).Gap = 15
                '                .Titles(0).URL = "."
                '                .Gallery = Gallery_Pie
                '                .Chart3D = False
                '                .OpenData COD_Values, 1, RsChartCredit.RecordCount
                '                For I = 0 To RsChartCredit.RecordCount - 1
                '                    If RsChartCredit("COl1").Value = 10000 Then
                '                        .Value(0, I) = RsChartCredit("XXX").Value
                '                        .Legend.Item(I) = "√ﬂÀ— „‰ 90 ÌÊ„"
                '                    Else
                '                        .Value(0, I) = RsChartCredit("XXX").Value
                '                        .Legend.Item(I) = "√ﬁ· „‰ " & RsChartCredit("COl1").Value & " ÌÊ„"
                '                    End If
                '                    RsChartCredit.MoveNext
                '                Next I
                '                .CloseData COD_Values
                '                '.PointLabels = True
                '                .ShowTips = True
                '                .LegendBox = True
                '                .GalleryObj.LabelsInside = True
                '                .GalleryObj.Shadows = True
                '            End With
                '        Else
                '            FrmPaymentTime.ChartRecv.ClearData ClearDataFlag_AllData
            End If
        End If

        If Not (RsDebitValues.BOF Or RsDebitValues.EOF) Then

            With FrmPaymentTime.Fg2
                .rows = .FixedRows + RsDebitValues.RecordCount
                RsDebitValues.MoveFirst

                For i = 1 To RsDebitValues.RecordCount
                    .TextMatrix(i, .ColIndex("TransactionsID")) = IIf(IsNull(RsDebitValues("TransactionsID").value), "", RsDebitValues("TransactionsID").value)
                    .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(RsDebitValues("Transaction_Type").value), "", RsDebitValues("Transaction_Type").value)
                    .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsDebitValues("TransactionTypeName").value), "", RsDebitValues("TransactionTypeName").value)
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDebitValues("CusName").value), "", RsDebitValues("CusName").value)
                    .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDebitValues("NoteID").value), "", RsDebitValues("NoteID").value)
                    .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(RsDebitValues("NoteType").value), "", RsDebitValues("NoteType").value)
                    .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDebitValues("Note_Value").value), "", RsDebitValues("Note_Value").value)

                    If Not IsNull(RsDebitValues("DueDate").value) Then
                        .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsDebitValues("DueDate").value)
                        .TextMatrix(i, .ColIndex("LateInterval")) = DateDiff("d", RsDebitValues("DueDate").value, Date)
                    End If

                    .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsDebitValues("Transaction_Serial").value), "", RsDebitValues("Transaction_Serial").value)

                    If Not IsNull(RsDebitValues("Transaction_Date").value) Then
                        .TextMatrix(i, .ColIndex("Transaction_Date")) = DisplayDate(RsDebitValues("Transaction_Date").value)
                    End If

                    .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsDebitValues("CusID").value), "", RsDebitValues("CusID").value)

                    If Not IsNull(RsDebitValues("NoteDate").value) Then
                        .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(RsDebitValues("NoteDate").value)
                    End If

                    .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(RsDebitValues("NotesTypeName").value), "", RsDebitValues("NotesTypeName").value)
                    .TextMatrix(i, .ColIndex("PreRelease")) = IIf(IsNull(RsDebitValues("PreRelease").value), "", RsDebitValues("PreRelease").value)
                    .TextMatrix(i, .ColIndex("RequiredValue")) = IIf(IsNull(RsDebitValues("RequiredValue").value), "", RsDebitValues("RequiredValue").value)
                    RsDebitValues.MoveNext
                Next i

                DrawFloodProgress FrmPaymentTime.Fg2, .ColIndex("LateInterval")
                .AutoSize 0, .Cols - 1, False
            End With

        End If

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            If Not (RsChartDebit.BOF Or RsChartDebit.EOF) Then
                '            With FrmPaymentTime.ChartPay
                '                .PointLabelMask = maskString
                '                .PointLabels = False
                '                .Titles(0).text = "„ «Œ—«  «·œ›⁄ ··„Ê—œÌ‰"
                '                .Titles(0).Alignment = StringAlignment_Near
                '                .Titles(0).LineAlignment = StringAlignment_Near
                '                .Titles(0).Font.Bold = False
                '                .Titles(0).Font.Charset = 178
                '                .Titles(0).DockArea = DockArea_Top
                '                .Titles(0).DrawingArea = True
                '                .Titles(0).Gap = 15
                '                .Titles(0).URL = "."
                '                .Gallery = Gallery_Pie
                '                .Chart3D = False
                '                .OpenData COD_Values, 1, RsChartDebit.RecordCount
                '                For I = 0 To RsChartDebit.RecordCount - 1
                '                    If RsChartDebit("COl1").Value = 10000 Then
                '                        .Value(0, I) = RsChartDebit("XXX").Value
                '                        .Legend.Item(I) = "√ﬂÀ— „‰ 90 ÌÊ„"
                '                    Else
                '                        .Value(0, I) = RsChartDebit("XXX").Value
                '                        .Legend.Item(I) = "√ﬁ· „‰ " & RsChartDebit("COl1").Value & " ÌÊ„"
                '                    End If
                '                    RsChartDebit.MoveNext
                '                Next I
                '                .CloseData COD_Values
                '                '.PointLabels = True
                '                .ShowTips = True
                '                .LegendBox = True
                '                .GalleryObj.LabelsInside = True
                '                .GalleryObj.Shadows = True
                '            End With
                '        Else
                '            FrmPaymentTime.ChartPay.ClearData ClearDataFlag_AllData
            End If
        End If
    End If

    With FrmPaymentTime
        StrCaption = "⁄œœ «·√Ê—«ﬁ «·„«·Ì…:" & .FG1.rows - 1
        StrCaption = StrCaption & CHR(13)
        StrCaption = StrCaption & "≈Ã„«·Ï ﬁÌ„… «·√Ê—«ﬁ:" & .FG1.Aggregate(flexSTSum, 0, .FG1.ColIndex("RequiredValue"), .FG1.rows, .FG1.ColIndex("RequiredValue"))
        .lbl(0).Caption = StrCaption
        StrCaption = "⁄œœ «·√Ê—«ﬁ «·„«·Ì… :" & .Fg2.rows - 1
        StrCaption = StrCaption & CHR(13)
        StrCaption = StrCaption & "≈Ã„«·Ï ﬁÌ„… «·√Ê—«ﬁ:" & .Fg2.Aggregate(flexSTSum, 0, .Fg2.ColIndex("RequiredValue"), .Fg2.rows, .Fg2.ColIndex("RequiredValue"))
        .lbl(1).Caption = StrCaption
    End With

    'FrmPaymentTime.ApplySetting
    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    FrmPaymentTime.show
    Exit Function
ErrTrap:
End Function

Public Function ShowInstallmentMustPay(Optional ShowMsg As Boolean = False) As Boolean
    On Error GoTo ErrTrap
    Dim Msg    As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim rs     As New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "select * From QestNotReceipted where  DueDate<=#" & SQLDate(Date) & "#"
        StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(Date) & "'"
        StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"
    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTemp.EOF Or RsTemp.BOF Then
        If ShowMsg = True Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "No premiums are required"
            Else
                Msg = "·«  ÊÃœ √Ì √ﬁ”«ÿ Õ«‰ Êﬁ  ”œ«œÂ«"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "" ' App.Title
        End If

        ShowInstallmentMustPay = False
    Else
        ShowInstallmentMustPay = True
    End If

    Exit Function
ErrTrap:
End Function

Public Function ShowRequest(Optional ShowMsg As Boolean = False) As Boolean
    On Error GoTo ErrTrap
    Dim Msg    As String
    Dim RsTemp As New ADODB.Recordset
    RsTemp.Open "RequestItems", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If RsTemp.EOF Or RsTemp.BOF Then
        If ShowMsg = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«  ÊÃœ √’‰«› »·€  Õœ «·ÿ·»"
            Else
                Msg = "No Items On Demand"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If

        ShowRequest = False
    Else
        ShowRequest = True
    End If

    Exit Function
ErrTrap:
End Function

Public Sub ShowItemReportDialog(LngItemID As Long)

End Sub

Public Sub SetItemReportsMenu(LngItemID As Long, _
                              Frm As Form, _
                              Optional StrItemSerial As String = "", _
                              Optional LngStoreID As Long = 0)

    mdifrmmain.MnuItemTools.Tag = ""
    mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
    mdifrmmain.MnuItemTools_ItemCart.Tag = ""
    mdifrmmain.MnuItemTools_ItemData.Tag = ""
    mdifrmmain.MnuItemTools_ItemQty.Tag = ""
    mdifrmmain.MnuItemTools_ItemCostTrans.Tag = ""
        
    mdifrmmain.MnuItemTools.Tag = LngItemID

    If StrItemSerial <> "" Then
        mdifrmmain.MnuItemTools_ItemSerial.Enabled = True
        mdifrmmain.MnuItemTools_ItemSerial.Tag = LngItemID & "-" & StrItemSerial
    Else
        mdifrmmain.MnuItemTools_ItemSerial.Enabled = False
        mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
    End If

    mdifrmmain.MnuItemTools_ItemCart.Tag = LngItemID & "-" & LngStoreID
    mdifrmmain.MnuItemTools_ItemQty.Tag = LngItemID
    mdifrmmain.MnuItemTools_ItemData.Tag = LngItemID
    mdifrmmain.MnuItemTools_ItemCostTrans.Tag = LngItemID

    Frm.PopupMenu mdifrmmain.MnuItemTools
End Sub

Public Function CheckDelAccount1(StrAccountCode As String) As Boolean
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    StrSQL = "SELECT * from Accounts where  Account_Code like'" & StrAccountCode & "a%'"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        CheckDelAccount1 = False
    Else
        CheckDelAccount1 = True
    End If

End Function

Public Function CheckDelAccount(StrAccountCode As String) As Boolean
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset

    Dim rs2    As New ADODB.Recordset

    StrSQL = "select * from DOUBLE_ENTREY_VOUCHERS1 where Account_Code='" & StrAccountCode & "'"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs2.RecordCount > 0 Then
        CheckDelAccount = False
        Exit Function

    End If

    StrSQL = "SELECT dbo.ACCOUNTS.Account_ID,dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.Account_Name," & "dbo.ACCOUNTS.last_account,dbo.ACCOUNTS.cannot_del,dbo.ACCOUNTS.Account_Serial," & "dbo.ACCOUNTS.BasicAccount "
    StrSQL = StrSQL + ",COUNT(dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID) AS CountX"
    StrSQL = StrSQL + ",SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS SumX"

    StrSQL = StrSQL + " FROM dbo.ACCOUNTS LEFT JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON " & "dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code "
    StrSQL = StrSQL + " Where  (dbo.ACCOUNTS.Account_Code='" & StrAccountCode & "'"
    StrSQL = StrSQL + " OR dbo.ACCOUNTS.Account_Code Like '" & StrAccountCode & "%')"
    StrSQL = StrSQL + " GROUP BY dbo.ACCOUNTS.Account_ID, dbo.ACCOUNTS.Account_Code," & "dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.last_account,dbo.ACCOUNTS.cannot_del," & "dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.BasicAccount"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        CheckDelAccount = True
    ElseIf rs("cannot_del").value = True Then
        CheckDelAccount = False
    ElseIf rs("BasicAccount").value = True Then
        CheckDelAccount = False
    ElseIf rs("CountX").value = 0 Then
        CheckDelAccount = True
    Else
        CheckDelAccount = False
    End If

End Function

Public Function GetTransactionTotalPeriod(IntTransactionType As Integer, _
                                          FromDate As Variant, _
                                          ToDate As Variant)
    ' ” Œœ„ Â–Â «·œ«·… ·„⁄—›… ≈Ã„«·Ï Õ—ﬂ«  „⁄Ì‰…
    '›Ï › —… “„‰Ì… „Õœœ…
    Dim StrSQL  As String
    Dim rs      As ADODB.Recordset
    Dim DblTemp As Double

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = StrSQL + " SELECT  SUM(TotalAfterTax) AS SumX "
        StrSQL = StrSQL + " FROM  dbo.QryTransactionsTotal() QryTransactionsTotal "
        StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Type=" & IntTransactionType & ""

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(ToDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True)
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = StrSQL + " SELECT  SUM(TotalAfterTax) AS SumX "
        StrSQL = StrSQL + " FROM   QryTransactionsTotal "
        StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Type=" & IntTransactionType & ""

        If Not IsNull(FromDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(CDate(FromDate), True)
        End If

        If Not IsNull(ToDate) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(CDate(ToDate), True)
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        DblTemp = IIf(IsNull(rs("SumX").value), 0, rs("SumX").value)
    Else
        DblTemp = 0
    End If

    GetTransactionTotalPeriod = DblTemp
End Function

Public Function ShowRelatedTransactions(LngTransID As Long, _
                                        IntType As Integer, _
                                        Optional ByRef SngValues As Single) As Boolean

    Dim StrSQL  As String
    Dim rs      As ADODB.Recordset
    Dim i       As Integer
    Dim Frm     As FrmShowTransNotes
    Dim RsTrans As ADODB.Recordset
    On Error GoTo hErr
    ShowRelatedTransactions = False

    If LngTransID = 0 Then
        Exit Function
    End If

    'IntType =Request Only
    'IntaaType=Show Notes
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type,dbo.TransactionTypes." & "TransactionTypeName, dbo.Transactions.PaymentType, dbo.Transactions.CusID, dbo.TblCustemers.CusName," & "dbo.Transactions.StoreID , dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName," & "dbo.Transactions.ReturnID, dbo.QryOneTransactionTotal(dbo.Transactions.Transaction_ID) AS Sumx"
        StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN " & "dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type " & "INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN " & "dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN dbo.TblUsers ON " & "dbo.Transactions.UserID = dbo.TblUsers.UserID "
        StrSQL = StrSQL + " Where dbo.Transactions.ReturnID=" & LngTransID
        StrSQL = StrSQL + " Order By dbo.Transactions.Transaction_ID "
    
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = " SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Date, Transactions.Transaction_Type, TransactionTypes.TransactionTypeName," & "Transactions.PaymentType, Transactions.CusID, TblCustemers.CusName, Transactions.StoreID, TblStore.StoreName," & "Transactions.UserID, TblUsers.UserName, Transactions.ReturnID, QryTransactionsTotal.TotalAfterTax as SumX"
        StrSQL = StrSQL + " FROM (TransactionTypes INNER JOIN (TblUsers INNER JOIN (TblStore INNER JOIN " & "(TblCustemers INNER JOIN Transactions ON TblCustemers.CusID = Transactions.CusID) ON TblStore.StoreID =" & "Transactions.StoreID) ON TblUsers.UserID = Transactions.UserID) ON TransactionTypes.Transaction_Type = " & "Transactions.Transaction_Type) INNER JOIN QryTransactionsTotal ON Transactions.Transaction_ID = " & "QryTransactionsTotal.Transaction_ID "
        StrSQL = StrSQL + " Where Transactions.ReturnID=" & LngTransID
        StrSQL = StrSQL + " Order By Transactions.Transaction_ID "
    
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.BOF Or rs.EOF) Then
        ShowRelatedTransactions = False
        rs.Close
        Set rs = Nothing
        Exit Function
    End If

    If IntType = 0 Then
        If Not IsMissing(SngValues) Then

            For i = 0 To rs.RecordCount - 1
                SngValues = SngValues + IIf(IsNull(rs("Sumx").value), 0, rs("Sumx").value)
                rs.MoveNext
            Next i

        End If

        rs.Close
        Set rs = Nothing
        ShowRelatedTransactions = True
        Exit Function
    End If

    '--------------------------------------------------------------------------------------------------
    Set Frm = New FrmShowTransNotes
    Frm.IntMode = 1

    If SystemOptions.UserInterface = ArabicInterface Then
        Frm.Caption = "«·Õ—ﬂ«  «· Ã«—Ì…"
        Frm.lbl(0).Caption = "»Ì«‰«  «·Õ—ﬂ«  «· Ã«—Ì… «·ﬁ«∆„… ⁄·Ï «·Õ—ﬂ…"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        Frm.Caption = "Transactions"
        Frm.lbl(0).Caption = "Retruned Transactions"
    End If

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusID,dbo.TblCustemers.CusName ," & "dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName "
        StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN  dbo.TblCustemers ON " & "dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN dbo.TransactionTypes ON " & "dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
        StrSQL = StrSQL + " Where  dbo.Transactions.Transaction_ID=" & LngTransID & ""
    
        Debug.Print Replace(StrSQL, "dbo.", "", , , vbTextCompare)
    
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = ""
        StrSQL = " SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Date,TblCustemers.CusID,TblCustemers.CusName ," & "Transactions.Transaction_Type,TransactionTypes.TransactionTypeName "
        StrSQL = StrSQL + " FROM (Transactions INNER JOIN  TblCustemers ON " & "Transactions.CusID = TblCustemers.CusID )INNER JOIN TransactionTypes ON " & "Transactions.Transaction_Type =TransactionTypes.Transaction_Type "
        StrSQL = StrSQL + " Where  Transactions.Transaction_ID=" & LngTransID & ""
    End If

    Set RsTrans = New ADODB.Recordset
    RsTrans.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTrans.BOF Or RsTrans.EOF) Then
        Frm.lbl(7).Caption = IIf(IsNull(RsTrans("Transaction_Serial").value), "", RsTrans("Transaction_Serial").value)
        Frm.lbl(9).Caption = IIf(IsNull(RsTrans("TransactionTypeName").value), "", RsTrans("TransactionTypeName").value)

        If Not IsNull(RsTrans("Transaction_Date").value) Then
            Frm.lbl(12).Caption = DisplayDate(RsTrans("Transaction_Date").value)
        End If

        Frm.LblLink.Caption = IIf(IsNull(RsTrans("CusName").value), "", RsTrans("CusName").value)
        Frm.LblLink.Tag = IIf(IsNull(RsTrans("CusID").value), "", RsTrans("CusID").value)
    End If

    RsTrans.Close
    Set RsTrans = Nothing
    '--------------------------------------------------------------------------------------------------

    With Frm.fg
        '------------------------------------
        .Cols = 0
        .Cols = 11
        .FixedRows = 1
        .ColKey(0) = "Serial"
        .ColKey(1) = "Transaction_ID"
        .ColHidden(1) = True
        .ColKey(2) = "TransactionTypeName"
        .ColKey(3) = "Transaction_Serial"
        .ColKey(4) = "Transaction_Date"
        .ColKey(5) = "PaymentType"
        .ColKey(6) = "CusName"
        .ColKey(7) = "StoreName"
        .ColKey(8) = "UserName"
        .ColKey(9) = "Sumx"
        .ColKey(10) = "Transaction_Type"
        .ColHidden(10) = True
        '------------------------------------
    
        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True
            .cell(flexcpAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignRightCenter
            .TextMatrix(0, .ColIndex("Serial")) = "„"
            .TextMatrix(0, .ColIndex("Transaction_ID")) = "—ﬁ„ «·Õ—ﬂ…"
            .TextMatrix(0, .ColIndex("TransactionTypeName")) = "‰Ê⁄ «·Õ—ﬂ…"
            .TextMatrix(0, .ColIndex("Transaction_Serial")) = "„”·”· «·Õ—ﬂ…"
            .TextMatrix(0, .ColIndex("Transaction_Date")) = " «—ÌŒ «·Õ—ﬂ…"
            .TextMatrix(0, .ColIndex("PaymentType")) = "ÿ—Ìﬁ… «·œ›⁄"
            .TextMatrix(0, .ColIndex("CusName")) = "«”„ «·⁄„Ì·"
            .TextMatrix(0, .ColIndex("StoreName")) = "«”„ «·„Œ“‰"
            .TextMatrix(0, .ColIndex("UserName")) = "«”„ «·„Õ——"
            .TextMatrix(0, .ColIndex("Sumx")) = "≈Ã„«·Ï «·Õ—ﬂ…"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .RightToLeft = False
            .cell(flexcpAlignment, 0, 0, .rows - 1, .Cols - 1) = flexAlignLeftCenter
            .TextMatrix(0, .ColIndex("Serial")) = "Serial"
            .TextMatrix(0, .ColIndex("Transaction_ID")) = "Transaction ID"
            .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type Name"
            .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Transaction Serial"
            .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction Date"
            .TextMatrix(0, .ColIndex("PaymentType")) = "Payment Type"
            .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
            .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
            .TextMatrix(0, .ColIndex("UserName")) = "User Name"
            .TextMatrix(0, .ColIndex("Sumx")) = "Total"
        End If

        .AutoSize 0, .Cols - 1, False
        .rows = .FixedRows + rs.RecordCount

        For i = 1 To .rows - 1
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
            .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)

            If Not (IsNull(rs("Transaction_Date").value)) Then
                .TextMatrix(i, .ColIndex("Transaction_Date")) = DisplayDate(rs("Transaction_Date").value)
            End If

            If Not IsNull(rs("PaymentType").value) Then
                If rs("PaymentType").value = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("PaymentType")) = "‰ﬁœÏ"
                    Else
                        .TextMatrix(i, .ColIndex("PaymentType")) = "Cash"
                    End If

                ElseIf rs("PaymentType").value = 1 Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("PaymentType")) = "√Ã·"
                    Else
                        .TextMatrix(i, .ColIndex("PaymentType")) = "Due"
                    End If
                End If
            End If

            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), 0, rs("StoreName").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
            .TextMatrix(i, .ColIndex("Sumx")) = IIf(IsNull(rs("Sumx").value), "", rs("Sumx").value)
            .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(rs("Transaction_Type").value), "", rs("Transaction_Type").value)
            rs.MoveNext
        Next i

        Frm.lbl(5).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("Transaction_ID"), .rows, .ColIndex("Transaction_ID"))
        Frm.lbl(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Sumx"), .rows, .ColIndex("Sumx"))
        .AutoSize 0, .Cols - 1, False
    End With

    Frm.show
    Exit Function
hErr:
    ShowRelatedTransactions = False
End Function

Private Sub UpdateDatabase()
    Dim StrSQLNewView As String
    Dim StrSQLOldView As String

    '3/5/2008
    DB_CreateField "Transaction_Details", "Remarks", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·’‰›", False, True
    '18/5/2008
    DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
    '9_5_2008
    StrSQLNewView = "SELECT Transaction_Details.Transaction_ID, Transactions.Transaction_Date," & "TblCustemers.CusName, TblStore.StoreName, Transaction_Details.Item_ID, TblItems.ItemName," & "TblItems.ItemCode, Transaction_Details.ItemCase, Transaction_Details.ItemSerial," & "Transaction_Details.Quantity, Transaction_Details.Price, Transaction_Details.ItemDiscountType," & "Transaction_Details.ItemDiscount, Transactions.Trans_Discount, Transactions.Trans_DiscountType," & "Transactions.TaxFound, Transactions.TaxValue, TblItems.HaveSerial, Transactions.Transaction_Serial," & "Transaction_Details.guaranteeTime, TblEmployee.Emp_Code, TblEmployee.Emp_Name, Transaction_Details.Remarks"
    StrSQLNewView = StrSQLNewView + " FROM (TblStore INNER JOIN (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN " & "Transactions ON TblCustemers.CusID = Transactions.CusID) ON TblEmployee.Emp_ID = Transactions.Emp_ID) " & "ON TblStore.StoreID = Transactions.StoreID) INNER JOIN (TblItems INNER JOIN Transaction_Details ON " & "TblItems.ItemID = Transaction_Details.Item_ID) ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "

    StrSQLOldView = "SELECT Transaction_Details.Transaction_ID, Transactions.Transaction_Date," & "TblCustemers.CusName, TblStore.StoreName, Transaction_Details.Item_ID, TblItems.ItemName," & "TblItems.ItemCode, Transaction_Details.ItemCase, Transaction_Details.ItemSerial," & "Transaction_Details.Quantity, Transaction_Details.Price, Transaction_Details.ItemDiscountType," & "Transaction_Details.ItemDiscount, Transactions.Trans_Discount, Transactions.Trans_DiscountType," & "Transactions.TaxFound, Transactions.TaxValue, TblItems.HaveSerial, Transactions.Transaction_Serial," & "Transaction_Details.guaranteeTime, TblEmployee.Emp_Code, TblEmployee.Emp_Name"
    StrSQLOldView = StrSQLOldView + " FROM (TblStore INNER JOIN (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN " & "Transactions ON TblCustemers.CusID = Transactions.CusID) ON TblEmployee.Emp_ID = Transactions.Emp_ID) " & "ON TblStore.StoreID = Transactions.StoreID) INNER JOIN (TblItems INNER JOIN Transaction_Details ON " & "TblItems.ItemID = Transaction_Details.Item_ID) ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "

    DB_UpDateView "QryBuyReport", StrSQLNewView

    StrSQLNewView = "SELECT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Date," & "TblCustemers.CusName, TblStore.StoreName, QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount, QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.PaymentType, QryTransactionsTotal.CusID, QryTransactionsTotal.TaxFound," & "QryTransactionsTotal.TaxValue, QryTransactionsTotal.TransSum, QryTransactionsTotal.TotalAfterTax," & "QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Type, TblBoxesData.BoxID," & "TblBoxesData.BoxName, Transactions.SaleType, QryTransactionsTotal.StoreID, QryTransactionsTotal.Emp_ID," & "TblEmployee.Emp_Name, Transactions.TransactionComment "
    StrSQLNewView = StrSQLNewView + " FROM ((((QryTransactionsTotal INNER JOIN TblCustemers ON " & "QryTransactionsTotal.CusID = TblCustemers.CusID) INNER JOIN TblStore ON QryTransactionsTotal.StoreID =" & "TblStore.StoreID) INNER JOIN Transactions ON QryTransactionsTotal.Transaction_ID = Transactions.Transaction_ID)" & "INNER JOIN TblEmployee ON QryTransactionsTotal.Emp_ID = TblEmployee.Emp_ID) LEFT JOIN (TblBoxesData RIGHT JOIN Notes ON " & "TblBoxesData.BoxID = Notes.BoxID) ON Transactions.Transaction_ID = Notes.Transaction_ID"
    StrSQLOldView = StrSQLOldView + " Where (((QryTransactionsTotal.Transaction_Type) = 2)) "

    DB_UpDateView "ReportSallingTime", StrSQLNewView
End Sub

Public Function EditTransStatus(LngTransID As Long, _
                                StrAction As String, _
                                ItemsGrid As ClsGrid) As Boolean
    Dim rs         As ADODB.Recordset
    Dim StrSQL     As String
    Dim i          As Long
    Dim LngFindRow As Long

    ' StrSQL = "SELECT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type,dbo.TransactionTypes.TransactionTypeName," & "dbo.Transactions.PaymentType, dbo.Transactions.CusID, dbo.Transactions.StoreID,dbo.Transactions.UserID," & "dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.ItemSerial," & "dbo.Transaction_Details.Quantity , dbo.Transaction_Details.Price ,dbo.Transaction_Details.[ID]," & "dbo.TblItems.ItemCode, dbo.TblItems.HaveSerial,dbo.TblItems.ItemName"

    ' StrSQL = StrSQL + " FROM dbo.Transaction_Details INNER JOIN " & "dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN " & "dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN " & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    '
    '    StrSQL = StrSQL + " Where Transactions.Transaction_ID=" & LngTransID & ""
    '    StrSQL = StrSQL + " Order BY  Transactions.Transaction_ID,Transaction_Details.[ID]"
    '    Set rs = New ADODB.Recordset
    '    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    ''
    '   If Not (rs.BOF Or rs.EOF) Then
    ''
    '        With FrmEditTransStatus.FgItems
    '''            .Rows = .FixedRows + rs.RecordCount
    '
    '            For i = 1 To .Rows - 1
    '                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("Item_ID").value), "", rs("Item_ID").value)
    '                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
    '                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
    '                .TextMatrix(i, .ColIndex("BeforeTransQty")) = GetItemStockToTrans(rs("Item_ID").value, LngTransID)
    '                .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
    '                LngFindRow = ItemsGrid.Grid.FindRow(rs("Item_ID").value, ItemsGrid.Grid.FixedRows, ItemsGrid.Grid.ColIndex("Name"), False, True)

    ''                If LngFindRow <> -1 Then
    '                   .TextMatrix(i, .ColIndex("EditQty")) = ItemsGrid.Grid.TextMatrix(LngFindRow, ItemsGrid.Grid.ColIndex("Count"))
    '        End If
    '
    ''                rs.MoveNext
    '           Next i
    '
    '            .AutoSize 0, .Cols - 1, False
    '        End With
    '
    '    End If

    '    FrmEditTransStatus.show vbModal
End Function

Public Function SQLVersion() As String

End Function

Public Sub ShowItemCostEffectForTrans(IntTransType As Integer, _
                                      Optional LngTransID As Long = 0, _
                                      Optional StrTransSerial As String = "")

End Sub

Public Function CheckItemInv(LngItemID As Long, _
                             StrItemSerial As String, _
                             LngTransID As Long) As Boolean
    Dim StrSQL As String
    Dim rs     As ADODB.Recordset
    Dim RsOpt  As ADODB.Recordset
    Set RsOpt = New ADODB.Recordset
    RsOpt.Open "select CheckSal from TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim myvar As Integer

    If RsOpt("CheckSal") Then myvar = 2 Else myvar = 21
    If StrItemSerial <> "" Then
        StrSQL = "select * From QryGuarantee where Item_ID=" & LngItemID
        StrSQL = StrSQL + " and ItemSerial='" & StrItemSerial & "'"
        StrSQL = StrSQL + " AND Transaction_ID='" & LngTransID & "'"
        StrSQL = StrSQL + " AND Transaction_Type=" & myvar
    Else
        StrSQL = "select * From QryGuarantee where Item_ID=" & LngItemID
        StrSQL = StrSQL + " AND Transaction_ID='" & LngTransID & "'"
        StrSQL = StrSQL + " AND Transaction_Type=" & myvar
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        CheckItemInv = False
    Else
        CheckItemInv = True
    End If

    rs.Close
    Set rs = Nothing
End Function

