Attribute VB_Name = "ModAccounts"
Option Explicit

Public DEFALUT_CURRENCY As String

Public DEFALUT_CURRENCY_DIV As String

Public DEFALUT_CURRENCYE As String

Public DEFALUT_CURRENCY_DIVE As String




'========================
' Õ”«» „” ÊÏ «·Õ”«» »«·«⁄ „«œ ⁄·Ï «·¬»«¡ («·√œﬁ)
' ÌÕ«›Ÿ ⁄·Ï ‰›” «”„ «·œ«·…: CountAs
'========================


' «·«› —«÷Ì: «·Õ”«» „‰ «·¬»«¡ ⁄»— ﬁ«⁄œ… «·»Ì«‰« 

Private Const ROOT_CODE As String = "r"   ' «⁄ »— "r" √Ê NULL/"" Ã–—
Private Const MAX_DEPTH As Long = 1000    ' Õ„«Ì… „‰ «·œÊ—« /«·⁄„ﬁ €Ì— «·„‰ÿﬁÌ

Private Const COUNTAS_DEFAULT_MODE As Long = 1   ' 1 = LevelMode_FromParents_DB

'======================== „”«⁄œ«  ========================

' ⁄œ¯ Õ—› a (·· Ê«›ﬁ „⁄ «·ﬂÊœ «·ﬁœÌ„)
' √Ê÷«⁄ «·Õ”«»
Public Enum AccountLevelMode
    LevelMode_OldByA = 0          ' «·ﬁœÌ„: ⁄œ¯ Õ—› a ›Ì Account_Code (·· Ê«›ﬁ ›ﬁÿ)
    LevelMode_FromParents_DB = 1  ' «·√œﬁ: „‰ ”·”·… «·¬»«¡ ⁄»— ﬁ«⁄œ… «·»Ì«‰« 
    LevelMode_FromParents_Cache = 2  ' «·√œﬁ + «·√”—⁄: „‰ ﬂ«‘ Dictionary(Account_Code -> Parent_Account_Code)
End Enum

Private Function LevelByA(ByVal Account_code As String) As Integer
    Dim i As Long, c As Integer
    For i = 1 To Len(Account_code)
        If mId$(Account_code, i, 1) = "a" Then c = c + 1
    Next
    LevelByA = c
End Function

' »œÌ· ··‹ Nz ›Ì VB6
Private Function NzVB6(ByVal v As Variant, ByVal Fallback As String) As String
    If IsNull(v) Then
        NzVB6 = Fallback
    Else
        NzVB6 = CStr(v)
    End If
End Function

'======================== DB ========================

' Ã·» «·√» „‰ ﬁ«⁄œ… «·»Ì«‰«  (Ãˆœ« »”Ìÿ∫ ⁄œ¯· «·«”„/«·”ﬂÌ„… ·Ê ·“„)
Private Function GetParentFromDB(ByVal Account_code As String, ByVal Cn As ADODB.Connection) As String
    Dim Cmd As ADODB.Command, rs As ADODB.Recordset
    Dim parentCOde As String

    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = Cn
        .CommandType = adCmdText
        .CommandText = "SELECT TOP 1 Parent_Account_Code " & _
                       "FROM dbo.ACCOUNTS WHERE Account_Code = ?"
        .Parameters.Append .CreateParameter("p1", adVarWChar, adParamInput, 100, Account_code)
    End With

    Set rs = Cmd.Execute
    If Not rs Is Nothing Then
        If Not rs.EOF Then parentCOde = Trim$(NzVB6(rs.Fields(0).value, ""))
        rs.Close
    End If
    Set rs = Nothing: Set Cmd = Nothing

    GetParentFromDB = parentCOde
End Function

' Õ”«» «·„” ÊÏ „‰ «·¬»«¡ ⁄»— ﬁ«⁄œ… «·»Ì«‰« 
Private Function LevelFromParents_DB(ByVal Account_code As String, ByVal Cn As ADODB.Connection) As Integer
    Dim seen As Object ' Scripting.Dictionary ·„‰⁄ «·œÊ—« 
    Dim cur As String, parentCOde As String
    Dim depth As Long, i As Long

    cur = Trim$(Account_code)
    If cur = "" Then LevelFromParents_DB = 0: Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    depth = 1
    seen.Add cur, True

    For i = 1 To MAX_DEPTH
        parentCOde = Trim$(GetParentFromDB(cur, Cn))

        If parentCOde = "" Or parentCOde = ROOT_CODE Then Exit For

        If seen.Exists(parentCOde) Then Exit For ' œÊ—…
        seen.Add parentCOde, True

        depth = depth + 1
        cur = parentCOde
    Next

    LevelFromParents_DB = depth
End Function

'======================== CACHE ========================

' Õ”«» «·„” ÊÏ „‰ «·¬»«¡ ⁄»— ﬂ«‘ Dictionary(Account_Code -> Parent_Account_Code)
Private Function LevelFromParents_Cache(ByVal Account_code As String, ByVal dictParents As Object) As Integer
    Dim seen As Object, cur As String, parentCOde As String
    Dim depth As Long, i As Long

    cur = Trim$(Account_code)
    If cur = "" Then LevelFromParents_Cache = 0: Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    depth = 1
    seen.Add cur, True

    For i = 1 To MAX_DEPTH
        If dictParents.Exists(cur) Then
            parentCOde = Trim$(CStr(dictParents(cur)))
        Else
            parentCOde = "" ' √» €Ì— „⁄—Ê› = Ã–—
        End If

        If parentCOde = "" Or LCase$(parentCOde) = LCase$(ROOT_CODE) Then Exit For
        If seen.Exists(parentCOde) Then Exit For ' œÊ—…

        seen.Add parentCOde, True
        depth = depth + 1
        cur = parentCOde
    Next

    LevelFromParents_Cache = depth
End Function

'  Õ„Ì· ﬂ«‘ «·¬»«¡ „—… Ê«Õœ… («Œ Ì«—Ì)
Public Function BuildParentsCache(ByVal Cn As ADODB.Connection) As Object
    Dim dict As Object, rs As ADODB.Recordset, sql As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    sql = "SELECT Account_Code, Parent_Account_Code FROM dbo.ACCOUNTS"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly

    Do While Not rs.EOF
        dict(Trim$(NzVB6(rs!Account_code, ""))) = Trim$(NzVB6(rs!Parent_Account_Code, ""))
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing

    Set BuildParentsCache = dict
End Function

'======================== «·Ê«ÃÂ… «·—∆Ì”Ì… (ÌÕ«›Ÿ ⁄·Ï «·«”„) ========================

' CountAs:
' - Account_Code: ≈Ã»«—Ì
' - Account_Serial: „ÊÃÊœ ·· Ê«›ﬁ ›ﬁÿ Ê·« Ìı” Œœ„ Â‰«
' - Mode: «› —«÷Ì = FromParents_DB (√œﬁ)
' - Cn: « ’«· ﬁ«⁄œ… «·»Ì«‰«  (·Ê Mode=DB)
' - dictParents: «·ﬁ«„Ê” (·Ê Mode=CACHE)
Public Function CountAs2(ByVal Account_code As String, _
                        Optional ByVal account_serial As Variant, _
                        Optional ByVal Mode As AccountLevelMode = COUNTAS_DEFAULT_MODE, _
                        Optional ByVal Cn As ADODB.Connection, _
                        Optional ByVal dictParents As Object) As Integer
    On Error GoTo eh

    Select Case Mode
        Case LevelMode_OldByA
            CountAs2 = LevelByA(Account_code)

        Case LevelMode_FromParents_DB
            If Cn Is Nothing Then
                CountAs2 = 0  ' „—¯— Cn
            Else
                CountAs2 = LevelFromParents_DB(Account_code, Cn)
            End If

        Case LevelMode_FromParents_Cache
            If dictParents Is Nothing Then
                CountAs2 = 0  ' „—¯— «·ﬁ«„Ê”
            Else
                CountAs2 = LevelFromParents_Cache(Account_code, dictParents)
            End If

        Case Else
            CountAs2 = LevelByA(Account_code)
    End Select
    Exit Function

eh:
    CountAs2 = 0
End Function



Public Function UPDATE_ACCOUNT_COST_CENTER(StrAccountCode As String, _
                                           cost_center As Boolean, _
                                           cost_center_type As Integer, _
                                           cost_center_id As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    UPDATE_ACCOUNT_COST_CENTER = True
    StrSQL = "Select * From Accounts Where Account_Code='" & StrAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If Not (rs.BOF Or rs.EOF) Then
   
        rs("cost_center").value = cost_center
      
        rs("cost_center_type").value = cost_center_type
        rs("cost_center_id").value = cost_center_id
           
        rs.update
    Else
        UPDATE_ACCOUNT_COST_CENTER = False
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function GET_DEFAULT_CURRENCY_INF(Optional ID As Integer = 0) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    GET_DEFAULT_CURRENCY_INF = True
    
  If ID = 0 Then
    StrSQL = "Select * From currency Where basic=1"
  Else
  StrSQL = "Select * From currency Where id=" & ID
  End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        DEFALUT_CURRENCY = IIf(IsNull(rs("NAME").value), "", rs("NAME").value)
        DEFALUT_CURRENCYE = IIf(IsNull(rs("NAMEE").value), "", rs("NAMEE").value)

        DEFALUT_CURRENCY_DIV = IIf(IsNull(rs("divname").value), "", rs("divname").value)
        DEFALUT_CURRENCY_DIVE = IIf(IsNull(rs("divnameE").value), "", rs("divnameE").value)
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ ⁄„·… «· ⁄«„· «·«› —«÷Ì…"
        Else
            MsgBox "Please Define the default Currency"
        End If

        GET_DEFAULT_CURRENCY_INF = False
    End If

End Function

Public Function Get_Account_Serial(AccCode As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Code='" & AccCode & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Account_Serial = ""
        Exit Function
    End If
    If IsNull(Rs3("Account_Serial").value) Then
        Get_Account_Serial = ""
        Exit Function
    End If
    If Not IsNull(Rs3("Account_Serial").value) Then
        Get_Account_Serial = Rs3("Account_Serial").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function Get_Account_Name(Optional serial As String, _
                                 Optional Account_code As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Serial='" & serial & "'"

    If Account_code <> "" Then
        sql = "Select * from ACCOUNTS where Account_Code='" & Account_code & "'"
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Account_Name = ""
        Exit Function
    End If
    If IsNull(Rs3("Account_Name").value) And IsNull(Rs3("Account_NameEng").value) Then
        Get_Account_Name = ""
        Exit Function
    End If
  
    If SystemOptions.UserInterface = EnglishInterface Then
        If Not IsNull(Rs3("Account_NameEng").value) Then Get_Account_Name = Rs3("Account_NameEng").value
        Exit Function
    Else

        If Not IsNull(Rs3("Account_Name").value) Then Get_Account_Name = Rs3("Account_Name").value
        Exit Function
    End If
  
    Rs3.Close

End Function

Public Function Get_Name(tablename As String, _
                         Filedname As String, _
                         StringType As Boolean, _
                         Filedvalue As String, _
                         ByRef returnfiled As String, _
                         returnvalue As String)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from  " & tablename & "  where " & Filedname & "='" & Filedvalue & "'"

    If StringType = False Then
        sql = "Select * from  " & tablename & "  where " & Filedname & "=" & val(Filedvalue) & ""
    End If

    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        returnvalue = ""
        Exit Function
    End If
    If IsNull(Rs3(returnfiled).value) Then
        returnvalue = ""
        Exit Function
    End If
    If Not IsNull(Rs3(returnfiled).value) Then
        returnvalue = Rs3(returnfiled).value
        Exit Function
    End If
    Rs3.Close

End Function
Public Function Get_Account_code(serial As String, _
                                 Optional last_account As Integer = 0) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Serial='" & serial & "'"

    If last_account <> 0 Then
        sql = sql + " and  last_account=1"
    End If
    sql = sql + GetAccountByBarnchUser
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Account_code = ""
        Exit Function
    End If
    If IsNull(Rs3("Account_Code").value) Then
        Get_Account_code = ""
        Exit Function
    End If
    If Not IsNull(Rs3("Account_Code").value) Then
        Get_Account_code = Rs3("Account_Code").value
        Exit Function
    End If
    Rs3.Close

End Function
Public Function Get_Account_codex17072017(serial As String, _
                                 Optional last_account As Integer = 0) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Serial='" & serial & "'"

    If last_account <> 0 Then
        sql = sql + " and  last_account=1"
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    'If Rs3.RecordCount = 0 Then Get_Account_code = "": Exit Function
    'If IsNull(Rs3("Account_Code").value) Then Get_Account_code = "": Exit Function
    'If Not IsNull(Rs3("Account_Code").value) Then Get_Account_code = Rs3("Account_Code").value: Exit Function
    Rs3.Close

End Function
Public Function Get_Employee_Nationality(Name As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from Nationality where name='" & Name & "' or namee='" & Name & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Employee_Nationality = ""
        Exit Function
    End If
    If IsNull(Rs3("id").value) Then
        Get_Employee_Nationality = ""
        Exit Function
    End If
    If Not IsNull(Rs3("id").value) Then
        Get_Employee_Nationality = Rs3("id").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function Get_Employee_religon(Name As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
     
    sql = "Select * from dean where name='" & Name & "' or namee='" & Name & "'"
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Employee_religon = ""
        Exit Function
    End If
    If IsNull(Rs3("id").value) Then
        Get_Employee_religon = ""
        Exit Function
    End If
    If Not IsNull(Rs3("id").value) Then
        Get_Employee_religon = Rs3("id").value
        Exit Function
    End If
    Rs3.Close

End Function

Public Function Get_Account_Parent_code(Account_code As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Code='" & Account_code & "'"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Get_Account_Parent_code = ""
        Exit Function
    End If
    If IsNull(Rs3("Parent_Account_Code").value) Then
        Get_Account_Parent_code = ""
        Exit Function
    End If
    If Not IsNull(Rs3("Parent_Account_Code").value) Then
        Get_Account_Parent_code = Rs3("Parent_Account_Code").value
        Exit Function
    End If
    Rs3.Close

End Function


'Wael Return

Public Function get_account_max(account_serial As String, _
                                Optional StrParentAccCode As String) As Double
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String

    Dim max_no As String
    Dim i As Integer
    Dim ACCOUNT_CODE_AS As Integer
    ACCOUNT_CODE_AS = CountA(StrParentAccCode) + 1
    'Sql = "Select * from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    ' Sql = "Select max(cast(Account_Serial as float))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    '  Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    Dim account_root_lenght As Integer
    Dim max_no_lenght As Double
  
    'max_no_lenght = IIf(IsNull(Rs3("max_no").value) = False, Len(Rs3("max_no").value), 0)
  
    account_root_lenght = Len(account_serial)
 
    'Sql = "Select max(cast(right(Account_Serial , " & max_no_lenght - account_root_lenght & ") as float ))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'" & "AND LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & 1 'ACCOUNT_CODE_AS
    If SystemOptions.SuppCreat4Acc = True Then
        sql = "Select max(account_serial  )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where "
        'sql = sql & " account_serial NOT LIKE '%[a-zA-Z]%'"
        'sql = sql & " AND account_serial NOT LIKE '%[^0-9]%' and"
        sql = sql & " Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
    Else
    
        sql = "Select max(cast(account_serial as float) )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
    End If
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    Dim max_lenght As Double
 
    If Rs4.RecordCount = 0 Or IsNull(Rs4("max_no").value) Then
            get_account_max = "0"
            Exit Function
       End If
    Dim start_zero  As Integer
    start_zero = 0
    start_zero = 0

    If IsNull(Rs4("max_no1").value) Then
   
    Else

        For i = 1 To Len(Rs4("max_no1").value)

            If mId(Rs4("max_no1").value, i, 1) = "0" Then
                start_zero = start_zero + 1
                Else: GoTo mm
            End If
                    
        Next i

    End If

mm:
    max_no = IIf(IsNull(Rs4("max_no").value), 0, Rs4("max_no").value)
   If SystemOptions.SuppCreat4Acc = True Then
        max_no = ExtractNumberAfterCharacter(max_no)
   Else
   End If
   If max_no = "" Then
        max_no = Rs4("max_no").value & ""
   End If
    max_lenght = Len(max_no) - account_root_lenght + start_zero

    If max_lenght <= 0 Then GoTo ll
    max_no = right(max_no, max_lenght)
   
ll:
    get_account_max = val(max_no)

End Function


Public Function get_account_maxChar(account_serial As String, _
                                Optional StrParentAccCode As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String

    Dim max_no As String
    Dim i As Integer
    Dim ACCOUNT_CODE_AS As Integer
    ACCOUNT_CODE_AS = CountA(StrParentAccCode) + 1
    'Sql = "Select * from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    ' Sql = "Select max(cast(Account_Serial as float))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
    '  Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    Dim account_root_lenght As Integer
    Dim max_no_lenght As Double
  
    'max_no_lenght = IIf(IsNull(Rs3("max_no").value) = False, Len(Rs3("max_no").value), 0)
  
    account_root_lenght = Len(account_serial)
 
    'Sql = "Select max(cast(right(Account_Serial , " & max_no_lenght - account_root_lenght & ") as float ))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'" & "AND LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & 1 'ACCOUNT_CODE_AS
    If SystemOptions.SuppCreat4Acc = True Then
        sql = "Select max(account_serial  )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where "
        'sql = sql & " account_serial NOT LIKE '%[a-zA-Z]%'"
        'sql = sql & " AND account_serial NOT LIKE '%[^0-9]%' and"
        sql = sql & " Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
    Else
    
        sql = "Select max(cast(account_serial as float) )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
    End If
    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    Dim max_lenght As Double
 
    If Rs4.RecordCount = 0 Or IsNull(Rs4("max_no").value) Then
            get_account_maxChar = ""
            Exit Function
       End If
    Dim start_zero  As Integer
    start_zero = 0
    start_zero = 0

    If IsNull(Rs4("max_no1").value) Then
        
    Else
        
        'get_account_maxChar = ""
    End If
    
    max_no = ExtractCharacter(Rs4("max_no").value & "")
    
    get_account_maxChar = max_no

End Function

Function ExtractNumberAfterCharacter(inputString As String) As String
    Dim charIndex As Integer
    Dim i As Integer

    ' «·»ÕÀ ⁄‰ «·Õ—› ›Ì «·”·”·…
    For i = 1 To Len(inputString)
        If IsAlpha(mId(inputString, i, 1)) Then
            charIndex = i
            Exit For
        End If
    Next i

    ' ≈–«  „ «·⁄ÀÊ— ⁄·Ï «·Õ—›
    If charIndex > 0 Then
        ' Ì” Œ—Ã «·Ã“¡ »⁄œ «·Õ—›
        Dim numberPart As String
        numberPart = mId(inputString, charIndex + 1)

        ' ≈“«·… «·√Õ—› €Ì— «·—ﬁ„Ì…
        Dim numericPart As String
        numericPart = ""
        For i = 1 To Len(numberPart)
            If IsNumeric22(mId(numberPart, i, 1)) Then
                numericPart = numericPart & mId(numberPart, i, 1)
            End If
        Next i

        ExtractNumberAfterCharacter = numericPart
    Else
        ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·Õ—›
        ExtractNumberAfterCharacter = ""
    End If
End Function


Function ExtractCharacter(inputString As String) As String
    Dim charIndex As Integer
    Dim i As Integer

    ' «·»ÕÀ ⁄‰ «·Õ—› ›Ì «·”·”·…
    For i = 1 To Len(inputString)
        If IsAlpha(mId(inputString, i, 1)) Then
            charIndex = i
            Exit For
        End If
    Next i

    ' ≈–«  „ «·⁄ÀÊ— ⁄·Ï «·Õ—›
    If charIndex > 0 Then
        ' Ì” Œ—Ã «·Õ—›
        ExtractCharacter = mId(inputString, charIndex, 1)
    Else
        ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·Õ—›
        ExtractCharacter = ""
    End If
End Function



Function IsAlpha(s As String) As Boolean
    ' ÌÕœœ „« ≈–« ﬂ«‰ «·Õ—› √·›»«∆Ì«
    IsAlpha = (s >= "A" And s <= "Z") Or (s >= "a" And s <= "z")
End Function

Function IsNumeric22(s As String) As Boolean
    ' ÌÕœœ „« ≈–« ﬂ«‰ «·Õ—› —ﬁ„Ì«
    IsNumeric22 = s >= "0" And s <= "9"
End Function

Sub TestExtractNumberAfterCharacter()
    ' «” Œœ«„ «·›‰ﬂ‘‰
    Dim Result As String
    Result = ExtractNumberAfterCharacter("212f31228")
    MsgBox Result
End Sub
'
'Public Function get_account_max(account_serial As String, _
'                                Optional StrParentAccCode As String) As Double
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim Rs4 As ADODB.Recordset
'    Set Rs4 = New ADODB.Recordset
'
'    Dim sql As String
'
'    Dim max_no As String
'    Dim i As Integer
'    Dim ACCOUNT_CODE_AS As Integer
'    ACCOUNT_CODE_AS = CountA(StrParentAccCode) + 1
'    'Sql = "Select * from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
'    ' Sql = "Select max(cast(Account_Serial as float))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'"
'    '  Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    Dim account_root_lenght As Integer
'    Dim max_no_lenght As Double
'
'    'max_no_lenght = IIf(IsNull(Rs3("max_no").value) = False, Len(Rs3("max_no").value), 0)
'
'    account_root_lenght = Len(account_serial)
'
'    'Sql = "Select max(cast(right(Account_Serial , " & max_no_lenght - account_root_lenght & ") as float ))  as max_no from ACCOUNTS where Account_Serial like'" & account_serial & "%'" & "AND LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & 1 'ACCOUNT_CODE_AS
'    If SystemOptions.SuppCreat4Acc = True Then
'        sql = "Select max(account_serial  )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
'    Else
'
'        sql = "Select max(cast(account_serial as float) )  as max_no , max(account_serial   )  as max_no1  from ACCOUNTS where Account_Serial like'" & account_serial & "%' AND LEN(account_code) -LEN(REPLACE(ACCOUNT_CODE, 'a', ''))=" & ACCOUNT_CODE_AS
'    End If
'    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    Dim max_lenght As Double
'
'    If Rs4.RecordCount = 0 Or IsNull(Rs4("max_no").value) Then
'            get_account_max = 0
'            Exit Function
'       End If
'    Dim start_zero  As Integer
'    start_zero = 0
'    start_zero = 0
'
'    If IsNull(Rs4("max_no1").value) Then
'
'    Else
'
'        For i = 1 To Len(Rs4("max_no1").value)
'
'            If mId(Rs4("max_no1").value, i, 1) = "0" Then
'                start_zero = start_zero + 1
'                Else: GoTo mm
'            End If
'
'        Next i
'
'    End If
'
'mm:
'    max_no = IIf(IsNull(Rs4("max_no").value), 0, Rs4("max_no").value)
'
'    max_lenght = Len(max_no) - account_root_lenght + start_zero
'
'    If max_lenght <= 0 Then GoTo ll
'    max_no = right(max_no, max_lenght)
'
'll:
'    get_account_max = max_no
'
'End Function

'Public Function get_account_max_old(account_serial As String) As Integer
'Dim Rs3 As ADODB.Recordset
'Set Rs3 = New ADODB.Recordset
'Dim Sql As String
'Dim max As Integer
'Dim i As Integer
'
' Sql = "Select * from ACCOUNTS where Account_Serial like'" & account_serial & "__'"
'
'  Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'  If Rs3.RecordCount = 0 Then get_account_max = 0: Exit Function
'   max = Val(right(Rs3("Account_Serial").value, 2))
'  For i = 0 To Rs3.RecordCount - 1
'   If max < Val(right(Rs3("Account_Serial").value, 2)) Then
'   max = Val(right(Rs3("Account_Serial").value, 2))
'   End If
'   Rs3.MoveNext
'  Next i
'
' get_account_max = max
'
'End Function
'
Public Function ParentAccountPrperties(StrParentAccCode As String, _
                                       Optional ByRef AccountTypes As Integer = 0, _
                                       Optional ByRef AccountTab As Integer = 0, _
                                       Optional ByRef DepitOrCreditv As Integer = 0, _
                                       Optional ByRef Differenttypev As Integer = 0, _
                                       Optional ByRef Authorityv As Integer = 0, _
                                       Optional ByRef UserGroupIdv As Integer = 0, _
                                       Optional ByRef UserIdv As Integer = 0)

    Dim StrSQL As String
    Dim rs As New ADODB.Recordset

    StrSQL = "Select * from ACCOUNTS where Account_Code='" & StrParentAccCode & "'"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        AccountTypes = rs("AccountTypes").value
        AccountTab = rs("AccountTab").value
        DepitOrCreditv = IIf(IsNull(rs("DepitOrCredit").value), 0, rs("DepitOrCredit").value)
        Differenttypev = IIf(IsNull(rs("Differenttype").value), 1, rs("Differenttype").value)
        Authorityv = rs("Authority").value

    Else

    End If

End Function

Public Function AddNewAccount(StrParentAccCode As String, _
                              StrNewAccountName As String, _
                              Optional BolLastAcc As Boolean = True, _
                              Optional BolCannotDel As Boolean = False, _
                              Optional StrNewAccountNamee As String = "", _
                              Optional currenct_code As String = 1, _
                              Optional budget As Boolean = False, _
                              Optional cost_center As Boolean = False, _
                              Optional Sum_account As Boolean = False, _
                              Optional Branch As String = "0", _
                              Optional serial As String, _
                              Optional cost_center_type As Integer = 0, _
                              Optional cost_center_id As String, _
                              Optional ActivityTypeId As Integer = 0, _
                              Optional AccountTypes As Integer = 0, _
                              Optional AccountTab As Integer = 0, _
                              Optional DepitOrCreditv As Integer = 0, _
                              Optional Differenttypev As Integer = 0, _
                              Optional Authorityv As Integer = 0, _
                              Optional UserGroupIdv As Integer = 0, _
                              Optional UserIdv As Integer = 0, _
                              Optional ChKBlock As Boolean = False, _
                              Optional account_serial1 As String, _
                              Optional account_name1 As String, _
                              Optional Account_NameEng1 As String, Optional account_serial2 As String, Optional account_name2 As String, Optional Account_NameEng2 As String, Optional account_serial3 As String, Optional account_name3 As String, Optional Account_NameEng3 As String, Optional TblLCID As Long = 0)
      
    ParentAccountPrperties StrParentAccCode, AccountTypes, AccountTab, DepitOrCreditv, Differenttypev, Authorityv

    If CHECK_LAST_ACCOUNT(StrParentAccCode) = True Then
        MsgBox "·«Ì„ﬂ‰ «‰‘«¡ Õ”«»  Õ  «·Õ”«» «·‰Â«∆Ì :  " & Get_Account_Serial(StrParentAccCode): AddNewAccount = ""
        Exit Function
    End If
    Dim StrSQL        As String
    Dim rs            As ADODB.Recordset
    Dim StrNewAccCode As String
    Dim VarTemp       As Variant
    Dim s As String
    Dim i             As Integer, j As Integer
 
    StrSQL = "SELECT  ACCOUNTS.ActivityTypeId , ACCOUNTS.Branch , ACCOUNTS.Sum_account ,ACCOUNTS.cost_center ,ACCOUNTS.mowazna,ACCOUNTS.currenct_code,ACCOUNTS.Account_ID,Account_Code,Account_Name,Parent_Account_Code,ACCOUNTS.cost_center_type,  ACCOUNTS.cost_center_id" & ",last_account,cannot_del,Account_Serial,BasicAccount,DateCreated,Account_NameEng "
    StrSQL = StrSQL + " From ACCOUNTS "
    StrSQL = StrSQL + " Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccCode & "'))"
    StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID"

    StrSQL = " select * "
    StrSQL = StrSQL + " From ACCOUNTS "
    StrSQL = StrSQL + " Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccCode & "'))"
    StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID"

    StrSQL = " select * "
    StrSQL = StrSQL + " From ACCOUNTS where Parent_Account_Code='-1' "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Rs.BOF Or Rs.EOF Then
    '    StrNewAccCode = StrParentAccCode & "a" & 1
    'Else
    '    Rs.MoveLast
    '    VarTemp = Split(Rs("Account_Code").Value, "a", , vbTextCompare)
    '    I = VarTemp(UBound(VarTemp))
    '    StrNewAccCode = StrParentAccCode & "a" & I + 1
    '
    'End If
    Dim Count_ACCOUNT_digit As Integer
    Dim NoOfAs              As Integer

    StrNewAccCode = GetNewAcountCode(StrParentAccCode)
    NoOfAs = CountAs(StrParentAccCode) + 1

    'Count_ACCOUNT_digit = GetAccountsLevel(NoOfAs)

    'If NoOfAs = 1 Or NoOfAs = 2 Then
    'Count_ACCOUNT_digit = 0
    'ElseIf NoOfAs = 3 And NoOfAs = 4 Then
    'Count_ACCOUNT_digit = 2
    'Else
    'Count_ACCOUNT_digit = SystemOptions.Count_ACCOUNT_digit ' GetSetting(StrAppRegPath, "Setting", "COUNT_ACCOUNT_digit", 0)
    'End If
    Count_ACCOUNT_digit = GetAccountsLevel(NoOfAs)
    If SystemOptions.CustCreat4Acc Then
        If NoOfAs = 5 Then
            s = "SELECT Max(LEN(a.Account_Serial)) as  ff FROM ACCOUNTS a WHERE a.Parent_Account_Code = '" & StrParentAccCode & "'"
            Dim rsTestAcc As New ADODB.Recordset
            Set rsTestAcc = New ADODB.Recordset
            rsTestAcc.Open s, Cn, adOpenKeyset, adLockReadOnly
            If Not rsTestAcc.EOF Then
                If val(rsTestAcc!ff & "") < 10 Then
                    Count_ACCOUNT_digit = 3
                End If
            End If
            
        End If
    End If
'Count_ACCOUNT_digit = 5
    rs.AddNew
    rs("AccountTypes").value = AccountTypes
    rs("AccountTab").value = AccountTab
    rs("DepitOrCredit").value = DepitOrCreditv
    rs("Differenttype").value = Differenttypev
    rs("Authority").value = Authorityv
    rs("UserGroupId").value = UserGroupIdv
    rs("Userid").value = UserIdv
    rs("Block").value = ChKBlock
    
    rs("Account_Code").value = StrNewAccCode
    rs("Account_Name").value = StrNewAccountName
    rs("Parent_Account_Code").value = StrParentAccCode
    rs("last_account").value = BolLastAcc
    rs("cannot_del").value = BolCannotDel
    rs("Branch").value = Branch

    rs("account_serial1").value = account_serial1
    rs("account_name1").value = account_name1
    rs("Account_NameEng1").value = Account_NameEng1

    rs("account_serial2").value = account_serial2
    rs("account_name2").value = account_name2
    rs("Account_NameEng2").value = Account_NameEng2

    rs("account_serial3").value = account_serial3
    rs("account_name3").value = account_name3
    rs("Account_NameEng3").value = Account_NameEng3
    rs("TblLCID").value = TblLCID
    
    If Branch <> "" Then
    
        If Len(Branch) = 1 Then Branch = "00" & Branch
        If Len(Branch) = 2 Then Branch = "0" & Branch
             'Wael Return
        If serial = "" Then
            'If BolLastAcc = False Then
            'rs("Account_Serial").value = branch & Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(2, "0"))
            'Else
            'rs("Account_Serial").value = branch & Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(COUNT_ACCOUNT_digit, "0"))
            'End If
            
            If BolLastAcc = False Then
                If SystemOptions.SuppCreat4Acc Then
            
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & get_account_maxChar(Get_Account_Serial(StrParentAccCode), StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                Else
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                End If
            Else
                If SystemOptions.SuppCreat4Acc Then
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & get_account_maxChar(Get_Account_Serial(StrParentAccCode), StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                Else
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                End If
            End If

        Else
            rs("Account_Serial").value = serial
        
        End If

    Else

        If serial = "" Then

            '   If get_account_max(Get_Account_Serial(StrParentAccCode)) >= 9 Then
            If BolLastAcc = False Then
                If SystemOptions.SuppCreat4Acc Then
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & get_account_maxChar(Get_Account_Serial(StrParentAccCode), StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                Else
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                End If
            Else
                If SystemOptions.SuppCreat4Acc Then
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & get_account_maxChar(Get_Account_Serial(StrParentAccCode), StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                Else
                    rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & Format(get_account_max(Get_Account_Serial(StrParentAccCode), StrParentAccCode) + 1, String(Count_ACCOUNT_digit, "0"))
                End If
            End If

            '   Else
            '        rs("Account_Serial").value = Get_Account_Serial(StrParentAccCode) & "00" & get_account_max(Get_Account_Serial(StrParentAccCode)) + 1 ' Replace(StrNewAccCode, "a", "", , , vbTextCompare)
            '   End If
          
        Else
            rs("Account_Serial").value = serial
        End If
        
    End If
    
    rs("BasicAccount").value = False
    rs("DateCreated").value = Date

    If StrNewAccountNamee = "" Then
        rs("Account_NameEng").value = StrNewAccountName
    Else
        rs("Account_NameEng").value = StrNewAccountNamee
    End If
   
    rs("currenct_code").value = currenct_code
    rs("mowazna").value = budget
    rs("cost_center").value = cost_center
    rs("Sum_account").value = Sum_account
   
    rs("cost_center_type").value = cost_center_type
    rs("cost_center_id").value = cost_center_id
    rs("ActivityTypeId").value = ActivityTypeId
    
    rs.update
    rs.Close
    Set rs = Nothing
    AddNewAccount = StrNewAccCode
    Exit Function
End Function

Public Function GetAccountsLevel(AccountsLevelsid) As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select * from AccountsLevelsDetails where " & "Level" & "=" & AccountsLevelsid
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        GetAccountsLevel = 1
    Else
        GetAccountsLevel = IIf(IsNull(rs("NoOfDigits").value), 0, rs("NoOfDigits").value)
    End If

End Function

Public Function AddNewAccountold(StrParentAccCode As String, _
                                 StrNewAccountName As String, _
                                 Optional BolLastAcc As Boolean = True, _
                                 Optional BolCannotDel As Boolean = False, _
                                 Optional StrNewAccountNamee As String = "", _
                                 Optional currenct_code As String = 1, _
                                 Optional budget As Boolean = False, _
                                 Optional cost_center As Boolean = False, _
                                 Optional Sum_account As Boolean = False, _
                                 Optional Branch As String, _
                                 Optional serial As String)

    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim StrNewAccCode As String
    Dim VarTemp As Variant

    Dim i As Integer, j As Integer

    StrSQL = "SELECT ACCOUNTS.Branch , ACCOUNTS.Sum_account ,ACCOUNTS.cost_center ,ACCOUNTS.mowazna,ACCOUNTS.currenct_code,ACCOUNTS.Account_ID,Account_Code,Account_Name,Parent_Account_Code," & "last_account,cannot_del,Account_Serial,BasicAccount,DateCreated,Account_NameEng "
    StrSQL = StrSQL + " From ACCOUNTS "
    StrSQL = StrSQL + " Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccCode & "'))"
    StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'If Rs.BOF Or Rs.EOF Then
    '    StrNewAccCode = StrParentAccCode & "a" & 1
    'Else
    '    Rs.MoveLast
    '    VarTemp = Split(Rs("Account_Code").Value, "a", , vbTextCompare)
    '    I = VarTemp(UBound(VarTemp))
    '    StrNewAccCode = StrParentAccCode & "a" & I + 1
    '
    'End If
    StrNewAccCode = GetNewAcountCode(StrParentAccCode)
    rs.AddNew
    rs("Account_Code").value = StrNewAccCode
    rs("Account_Name").value = StrNewAccountName
    rs("Parent_Account_Code").value = StrParentAccCode
    rs("last_account").value = BolLastAcc
    rs("cannot_del").value = BolCannotDel

    If Branch <> "0" Then
    
        If Len(Branch) = 1 Then Branch = "00" & Branch
        If Len(Branch) = 2 Then Branch = "0" & Branch
        
        If serial = "" Then
            rs("Account_Serial").value = Branch & Replace(StrNewAccCode, "a", "", , , vbTextCompare)
        Else
            rs("Account_Serial").value = serial
        
        End If

    Else

        If serial = "" Then
            rs("Account_Serial").value = Replace(StrNewAccCode, "a", "", , , vbTextCompare)
        Else
            rs("Account_Serial").value = serial
        End If
        
    End If
    
    rs("BasicAccount").value = False
    rs("DateCreated").value = Date

    If StrNewAccountNamee = "" Then
        rs("Account_NameEng").value = StrNewAccountName
    Else
        rs("Account_NameEng").value = StrNewAccountNamee
    End If
   
    rs("currenct_code").value = currenct_code
    rs("mowazna").value = budget
    rs("cost_center").value = cost_center
    rs("Sum_account").value = Sum_account
    rs("Branch").value = Branch
    
    rs.update
    rs.Close
    Set rs = Nothing
    AddNewAccountold = StrNewAccCode
    Exit Function
End Function

Public Function check_account_exist(StrAccountCode As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From Accounts Where Account_Code='" & StrAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        check_account_exist = True
    Else
        check_account_exist = False
    End If

End Function

Public Function EditAccount(StrAccountCode As String, _
                            StrNewAccountName As String, _
                            Optional NameE As String, _
                            Optional mowazna As Boolean, _
                            Optional cost_center As Boolean, _
                            Optional currency_code As Integer = 1, _
                            Optional Sum_account As Boolean = 0, _
                            Optional serial As String, _
                            Optional cost_center_type As Integer = 0, _
                            Optional cost_center_id As String, _
                            Optional ActivityTypeId As Integer = 0, _
                            Optional AccountTypes As Integer = 0, _
                            Optional AccountTab As Integer = 0, _
                            Optional DepitOrCreditv As Integer = 0, _
                            Optional Differenttypev As Integer = 0, _
                            Optional Authorityv As Integer = 0, _
                            Optional UserGroupIdv As Integer = 0, _
                            Optional UserIdv As Integer = 0, _
                            Optional ChKBlock As Boolean = False, _
                            Optional BolLastAcc As Boolean = 0, Optional TblLCID As Long = 0)
 
    If AccountTypes = 0 Then
        ParentAccountPrperties StrAccountCode, AccountTypes, AccountTab, DepitOrCreditv, Differenttypev, Authorityv
    End If

    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From Accounts Where Account_Code='" & StrAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs("Account_Name").value = StrNewAccountName
        rs("Account_NameEng").value = NameE
        rs("mowazna").value = mowazna
        rs("cost_center").value = cost_center
        rs("currenct_code").value = currency_code
        rs("Sum_account").value = Sum_account
        rs("cost_center_type").value = cost_center_type
        rs("cost_center_id").value = cost_center_id
        '      If (BolLastAcc) <> 0 Then
        rs("last_account").value = BolLastAcc

        '    End If
        If serial <> "" Then
            rs("Account_Serial").value = serial
        End If

        rs("ActivityTypeId").value = ActivityTypeId
             
        rs("AccountTypes").value = AccountTypes
        rs("AccountTab").value = AccountTab
        rs("DepitOrCredit").value = DepitOrCreditv
        rs("Differenttype").value = Differenttypev
        rs("Authority").value = Authorityv
        rs("UserGroupId").value = UserGroupIdv
        rs("Userid").value = UserIdv
        rs("Block").value = ChKBlock
        rs("TblLCID").value = TblLCID
        
           
        rs.update
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function CanCreateNewInterval() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From TblAccountIntervals Where OpenState=1"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        'Data
        CanCreateNewInterval = False
    Else
        'NO Data
        CanCreateNewInterval = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetCurrentAccountIntervalID() As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    On Local Error GoTo ErrTrap
    StrSQL = "Select * From TblAccountIntervals Where OpenState=1"
    StrSQL = StrSQL + " Order BY AccountIntervalID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
        GetCurrentAccountIntervalID = rs("AccountIntervalID").value
    Else
        GetCurrentAccountIntervalID = 0
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    GetCurrentAccountIntervalID = 0
End Function

Public Function GetPaymentTypeBank(PaymentID As Long) As Long

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String
 
    StrSQL = "Select *  From  TblPaymentType "
    StrSQL = StrSQL + " Where PaymentID=  " & PaymentID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetPaymentTypeBank = IIf(IsNull(rs("bankid").value), 0, rs("bankid").value)
    Else
        GetPaymentTypeBank = 0
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetMyAccountCodeRefined(StrTableName, _
                                        StrIDFieldName As String, _
                                        Lngid As Long, _
                                        Optional Account_CodeFiled As String) As String

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Exit Function
    End If

    StrSQL = "Select " & Account_CodeFiled & " From " & StrTableName
    StrSQL = StrSQL + " Where " & StrIDFieldName & "=" & Lngid
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetMyAccountCodeRefined = IIf(IsNull(rs(Account_CodeFiled).value), "", rs(Account_CodeFiled).value)
    Else
        GetMyAccountCodeRefined = ""
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetMyAccountCode2(StrTableName, _
                                  StrIDFieldName As String, _
                                  Lngid As String) As String

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Exit Function
    End If

    StrSQL = "Select Account_Code From " & StrTableName
    StrSQL = StrSQL + " Where " & StrIDFieldName & "='" & Lngid & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetMyAccountCode2 = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
    Else
        GetMyAccountCode2 = ""
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetDriverAccountCode(DriverID As Long, Optional ByRef empSalaryAccount) As String

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Exit Function
    End If

    StrSQL = " SELECT dbo.TblEmployee.Account_Code1  as empSalaryAccount ,   dbo.TblBoxesData.Account_Code, dbo.TblBoxesData.ChequeBox, dbo.TblBoxesData.DriverId"
    StrSQL = StrSQL & "  FROM         dbo.TblBoxesData INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblBoxesData.empid = dbo.TblEmployee.Emp_ID"
    StrSQL = StrSQL & "   Where (dbo.TblBoxesData.type = 1) And (dbo.TblBoxesData.empid = " & DriverID & ")"
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetDriverAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
    '    empSalaryAccount = IIf(IsNull(rs("empSalaryAccount").value), "", rs("empSalaryAccount").value)
        
    Else
        GetDriverAccountCode = ""
    '    empSalaryAccount = IIf(IsNull(rs("empSalaryAccount").value), "", rs("empSalaryAccount").value)
        
    End If

    rs.Close
    Set rs = Nothing
End Function
Public Function GetemployeeAccountCode(Emp_id As Long) As String

    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

 

    StrSQL = " SELECT     Emp_ID, Account_code1"
StrSQL = StrSQL & "  From dbo.TblEmployee"
StrSQL = StrSQL & "   Where (emp_id = " & Emp_id & ")"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetemployeeAccountCode = IIf(IsNull(rs("Account_code1").value), "", rs("Account_code1").value)
     
        
    Else
        GetemployeeAccountCode = ""
     
        
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function GetSumOfGeForOneAccount(Account_code As String, _
                                 Transaction_ID As Double, _
                                 Optional Credit_Or_Debit As Integer = -1) As String
                                 
                                 
    Dim rs As ADODB.Recordset
    Dim StrSQL  As String

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Exit Function
    End If

    StrSQL = "   SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result"
    StrSQL = StrSQL & " From"
    StrSQL = StrSQL & " ("
     StrSQL = StrSQL & " SELECT"
     StrSQL = StrSQL & " DEV_Value1=Case"
     StrSQL = StrSQL & " When Credit_Or_Debit=0   Then Value * 1"
     StrSQL = StrSQL & " Else 0"
     StrSQL = StrSQL & " END,"
     StrSQL = StrSQL & " DEV_Value2=Case"
     StrSQL = StrSQL & " When Credit_Or_Debit=1  Then Value * 1"
     StrSQL = StrSQL & " Else 0"
     StrSQL = StrSQL & " End"
     StrSQL = StrSQL & " From dbo.DOUBLE_ENTREY_VOUCHERS"
    StrSQL = StrSQL & " WHERE     (Account_Code = N'" & Account_code & "') AND (Transaction_ID = " & Transaction_ID & ") AND (Credit_Or_Debit = 0)"
    If Credit_Or_Debit <> -1 Then
    StrSQL = StrSQL & " AND (Credit_Or_Debit = 0)"
    End If
    StrSQL = StrSQL & "  )x"
 
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetSumOfGeForOneAccount = IIf(IsNull(rs("result").value), 0, rs("result").value)
    Else
        GetSumOfGeForOneAccount = 0
    End If

    rs.Close
    Set rs = Nothing
    

End Function
Public Function GetMyAccountCode(StrTableName, _
                                 StrIDFieldName As String, _
                                 Lngid As Long, _
                                 Optional FieldName As String = "") As String

    Dim rs     As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Exit Function
    End If
    If FieldName = "" Then
        StrSQL = "Select Account_Code From " & StrTableName
    Else
        StrSQL = "Select " & FieldName & " From " & StrTableName
    End If
    
    StrSQL = StrSQL + " Where " & StrIDFieldName & "=" & Lngid
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        '    GetMyAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        If FieldName = "" Then
            GetMyAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
        Else
 
            GetMyAccountCode = IIf(IsNull(rs(FieldName).value), "", rs(FieldName).value)
        End If

    Else
        GetMyAccountCode = ""
    End If

    rs.Close
    Set rs = Nothing
End Function

 
      Public Function AddNewDev(LngDevID As Variant, _
                          IntLineNO As Variant, _
                          StrAccountCode As String, _
                          SngValue As Variant, _
                          Credit_Or_Debit As Integer, _
                          Optional StrDes As String, _
                          Optional LngNoteID As Variant = 0, _
                          Optional LngReceiptID As Long = 0, _
                          Optional LngOperaID As Long = 0, _
                          Optional IntAccInterval As Long = 0, _
                          Optional RecordDate As Date, _
                          Optional LngUserID As Long = 0, _
                          Optional LngTransaction_ID As Long = 0, _
                          Optional StrDEV_Serial As String = "", _
                          Optional LngAdvancedID As Long = 0, _
                          Optional valuee As Variant = 0, _
                          Optional curr As String = "", _
                          Optional Rate As Long = 1, Optional ExpensesID As Double, Optional StrDese As String, Optional IntLineNO1 As Double, Optional notes_all As Double, Optional project_id As Integer, Optional opr_fullcode As String, _
                          Optional opening_balance As Boolean = False, Optional opening_balance_voucher_id As Double, Optional FixedassetId As Integer, Optional FixedAssetgroupid As Integer, Optional FixedAssetbranch_id As Integer, _
                          Optional branch_id As Integer = 1, Optional CarID As Double, Optional ShowQty1 As Double = 0, Optional showPrice1 As Double = 0, Optional showPrice2 As Double = 0, Optional Salaries1 As Double = 0, Optional Salaries2 As Double = 0, _
                          Optional Departementid As Double = 0, Optional NEmpid As Double, Optional ContNo As Integer, Optional Aqarid As Integer, Optional unittype As Integer, Optional unitno As Integer, Optional BillNo As String, Optional project_id1 As Integer, Optional pand_id As Integer, _
                          Optional oper_id As Integer, Optional Remarks2 As String, Optional hideline As Integer = 0, Optional ToTrans As Integer, Optional BankID As Integer = 0, Optional BoxID As Integer = 0, Optional StoreID As Integer = 0, Optional EmpID As Integer = 0, Optional CusID As Integer = 0, _
                          Optional Posted As Integer = 0, Optional FLgBranch As Integer, Optional OtherInformation As ClsGLOther, Optional ByVal mDueDate As String = "", Optional ByVal IsHidden As Boolean = False, Optional ByVal project_bill_no As Long = 0) As Boolean

 

    Dim RsDev As ADODB.Recordset
    Dim RsSerial As ADODB.Recordset
    Dim StrSQL As String
    Dim LngSerialCount As Long
    Dim DblValue As Double
 
    'On Local Error GoTo ErrTrap
    
    DblValue = val(Format(SngValue, "." & String(Abs(3), "#")))
    DblValue = val(Format(SngValue, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))

    If DblValue = 0 Then

        AddNewDev = True
        Exit Function
    End If
    If IsMissing(RecordDate) Then
        RecordDate = Date
    End If

    Set RsDev = New ADODB.Recordset

    If opening_balance = False Then
        '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenForwardOnly, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* from dbo.DOUBLE_ENTREY_VOUCHERS Where (Double_Entry_Vouchers_ID = -1)"

        RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Else
        check_opening_balance_notes
        ' RsDev.Open "DOUBLE_ENTREY_VOUCHERS1", Cn, adOpenForwardOnly, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS1.* from dbo.DOUBLE_ENTREY_VOUCHERS1 Where (Double_Entry_Vouchers_ID = -1)"
        RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    End If

    RsDev.AddNew
    If Posted = 1 Then
        RsDev("Posted").value = Posted
    Else
        RsDev("Posted").value = Null
    End If
    If OtherInformation Is Nothing Then
        GoTo d
    End If
    If OtherInformation.FlgVat = 1 Then
        RsDev("FlgVat").value = 1
    Else
        RsDev("FlgVat").value = Null
    End If
    RsDev("CurrRow").value = OtherInformation.CurrRow
    RsDev("Vatyo").value = OtherInformation.Vatyo
    RsDev("Vat").value = Round(OtherInformation.Vat, 2)
    RsDev("TotalValue").value = Round(OtherInformation.TotalValue, 2)
    
    RsDev("Unitss").value = Trim(OtherInformation.Unitss)
    RsDev("StrUnit").value = Trim(OtherInformation.StrUnit)
   RsDev("mType").value = val(OtherInformation.mType)
   RsDev("iqarid").value = val(OtherInformation.iqarid)
    RsDev("uintid").value = val(OtherInformation.uintid)
 
 
    
    RsDev("DescAccount").value = OtherInformation.DescAccount
    RsDev("IsExpens").value = OtherInformation.isExpens
    
    RsDev("IsHidden").value = IIf(IsHidden, 1, 0)
    
    RsDev("AccountCode2").value = OtherInformation.AccountCode2
    RsDev("SupplierID").value = OtherInformation.SupplierID
    RsDev("CusVATNO").value = OtherInformation.CusVATNO
    RsDev("SupplierName").value = OtherInformation.SupplierName
    RsDev("PriceTotal").value = OtherInformation.PriceTotal
    RsDev("Rate2").value = OtherInformation.Rate
    RsDev("NextAccount_Code").value = OtherInformation.NextAccount_Code
    ' RsDev("BillNo").value = OtherInformation.BillNo
d:
    If opening_balance = True Then
        RsDev("opening_balance_voucher_id").value = opening_balance_voucher_id
        RsDev("ShowQty1").value = ShowQty1
        RsDev("showPrice1").value = showPrice1
        RsDev("showPrice2").value = showPrice2
        RsDev("Salaries1").value = Salaries1
        RsDev("Salaries2").value = Salaries2
        RsDev("opening_balance_voucher_id").value = opening_balance_voucher_id
        
        

        RsDev("mType").value = unittype
        RsDev("iqarid").value = Aqarid
         RsDev("uintid").value = unitno
    End If
    
    If hideline = 0 Then
        RsDev("hideline").value = Null
    Else
        RsDev("hideline").value = hideline
    End If


    If hideline = 6 Then
        RsDev("IsHiddenInv").value = 1
    Else
        RsDev("IsHiddenInv").value = Null
    End If

    RsDev("ToTrans").value = ToTrans
RsDev("project_bill_no").value = project_bill_no
    'If opening_balance = False Then
    RsDev("Remarks2").value = Remarks2
    RsDev("projectid").value = project_id1
    RsDev("pandid").value = pand_id
    RsDev("operid").value = oper_id
    If FLgBranch = 1 Then
        RsDev("FLgBranch").value = FLgBranch
    Else
        RsDev("FLgBranch").value = Null
    End If

    'End If
    RsDev("branch_id").value = branch_id
    RsDev("Double_Entry_Vouchers_ID").value = LngDevID
    RsDev("DEV_ID_Line_No").value = IntLineNO
    RsDev("DEV_ID_Line_No1").value = IntLineNO1
    
    RsDev("Account_Code").value = StrAccountCode
    DblValue = val(Format(SngValue, "." & String(Abs(3), "#")))
    '    DBLValue = Round(SngValue, SystemOptions.SysDefCurrencyForamt)
    RsDev("Value").value = DblValue
    RsDev("valuee").value = valuee
        
    '   RsDev("Value").value = Round(RsDev("Value").value, SystemOptions.SysDefCurrencyForamt)
    '   RsDev("Value").value = Round(RsDev("Value").value, SystemOptions.SysDefCurrencyForamt)
      
    '     RsDev("ExpensesID").value = ExpensesID
     
    RsDev("currency").value = curr
    RsDev("rate").value = Rate
    RsDev("Credit_Or_Debit").value = Credit_Or_Debit
    RsDev("Double_Entry_Vouchers_Description").value = StrDes
    RsDev("Double_Entry_Vouchers_Descriptione").value = StrDese
    
    If LngNoteID = 0 Then
        RsDev("Notes_ID").value = Null
    Else
        RsDev("Notes_ID").value = LngNoteID
    End If
    
    '  If Branch_Id = 0 Then
    '     rsdev("branch_id").value = Null
    ' Else
    '     rsdev("branch_id").value = LngNoteID
    ' End If
    
    If LngReceiptID = 0 Then
        RsDev("ReceiptID").value = Null
    Else
        RsDev("ReceiptID").value = LngReceiptID
    End If

    If LngOperaID = 0 Then
        RsDev("OperaID").value = Null
    Else
        RsDev("OperaID").value = LngOperaID
    End If

    If IntAccInterval = 0 Then
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
    Else
        RsDev("Account_Interval_ID").value = IntAccInterval
    End If

    RsDev("RecordDate").value = RecordDate
    RsDev("RecordDateH").value = ToHijriDate(RecordDate)
     
    If LngUserID = 0 Then
        RsDev("UserID") = user_id
    Else
        RsDev("UserID") = LngUserID
    End If

    If LngTransaction_ID = 0 Then
        RsDev("Transaction_ID").value = Null
    Else
        RsDev("Transaction_ID").value = LngTransaction_ID
    End If

    If LngAdvancedID = 0 Then
        RsDev("AdvanceID").value = Null
    Else
        RsDev("AdvanceID").value = LngAdvancedID
    End If

    If StrDEV_Serial <> "" Then
        RsDev("DEV_Serial").value = StrDEV_Serial
    Else
        RsDev("DEV_Serial").value = GetNewDEV_Serial(RecordDate)
    End If
    
    RsDev("notes_all").value = val(notes_all)
    RsDev("project_id").value = project_id
    RsDev("opr_fullcode").value = opr_fullcode
    
    RsDev("FixedAssetId").value = FixedassetId
    RsDev("FixedAssetgroupid").value = FixedAssetgroupid
    RsDev("FixedAssetbranch_id").value = FixedAssetbranch_id
    RsDev("Departementid").value = Departementid
    RsDev("NEmpid").value = NEmpid
    If opening_balance = True Then
        RsDev("ContNo").value = ContNo
    End If
    If opening_balance = False Then
        RsDev("CarId").value = CarID
      
        RsDev("Billno").value = BillNo
             RsDev("Aqarid").value = Aqarid
        RsDev("unittype").value = unittype
        RsDev("unitno").value = unitno
        RsDev("uintid").value = unitno

    End If


    If LngNoteID = 1 And opening_balance = True Then
        updateAutoOpeningBalanceVoucherValuebyCharacttex
    
    End If
    If mDueDate <> "" Then
    
        RsDev("DueDate").value = CDate(mDueDate)
    Else
        RsDev("DueDate").value = RecordDate
    End If

    RsDev.update

    AddNewDev = True
    RsDev.Close
    Set RsDev = Nothing
    Exit Function
ErrTrap:
    AddNewDev = False

    If RsDev.EditMode <> adEditNone Then
        RsDev.CancelUpdate
    End If

    RsDev.Close
    Set RsDev = Nothing
End Function




Public Function updateAutoOpeningBalanceVoucherValuebyCharacttex()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim total As Double
    sql = "SELECT     SUM([Value]) AS Total from dbo.DOUBLE_ENTREY_VOUCHERS1 WHERE  Credit_Or_Debit=0 and    (Notes_ID = 1)"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
    Else
        total = IIf(IsNull(rs(total).value), 0, rs(total).value)
        sql = "update Notes1 set note_value_by_characters='" & WriteNo(Format(total, "0.00"), 0, True, ".") & "' where NoteID=1"
        Cn.Execute sql

    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function AddNewOpenBalance(LngCusID As Long, _
                                  dtpDate As Date) As Long
    Dim RsNotes As ADODB.Recordset
    On Local Error GoTo ErrTrap
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open "NOTES", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    RsNotes.AddNew
    RsNotes("NoteID").value = new_id("NOTES", "NoteID", "", False)
    RsNotes("NoteDate").value = dtpDate
    RsNotes("NoteType").value = 101
    RsNotes("Note_Value").value = 0
    RsNotes("CusID").value = LngCusID
    RsNotes("Member_ID").value = LngCusID
    RsNotes("UserID") = user_id
    RsNotes("Remark").value = "—’Ìœ ≈›  «ÕÌ"
    RsNotes("NotePosted").value = 0
    RsNotes.update
    AddNewOpenBalance = RsNotes("NoteID").value
    RsNotes.Close
    Set RsNotes = Nothing
    Exit Function
ErrTrap:
    AddNewOpenBalance = 0
End Function

Private Function find_a_pos(X As String) As Integer
    Dim pos As Integer
    Dim i As Integer

    For i = 1 To Len(X)

        If mId(X, i, 1) = "a" Then
            pos = i
        End If

    Next i

    find_a_pos = pos

End Function

Private Function AccountCodeExists(ByVal AccountCode As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "SELECT TOP 1 1 AS X FROM ACCOUNTS WHERE Account_Code = '" & AccountCode & "'"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    AccountCodeExists = Not rs.EOF

    rs.Close
    Set rs = Nothing
End Function
Private Function GetNewAcountCode(StrParentAccountCode As String) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngMax As Long
    Dim pos As Integer
    Dim suffix As Long
    Dim NewCode As String

    ' √Ê·«: ‰ÃÌ» √ﬂ»— ”Ì—Ì«· ··√» œÂ „‰ √Ê·«œÂ
    StrSQL = "SELECT Account_Code " & _
             "FROM ACCOUNTS " & _
             "WHERE Parent_Account_Code = '" & StrParentAccountCode & "' " & _
             "ORDER BY Account_ID"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then
        ' „«›Ì‘ √»‰«¡ Œ«·’, ‰»œ√ „‰ 1
        suffix = 1
    Else
        LngMax = 0
        rs.MoveFirst

        Do While Not rs.EOF
            pos = find_a_pos(rs!Account_code)   ' ·«“„ Ì—Ã¯⁄ ¬Œ— a ›Ì «·ﬂÊœ

            If pos > 0 Then
                ' «·Ã“¡ «··Ì »⁄œ «·‹ a ÂÊ «·”Ì—Ì«·
                suffix = val(mId$(rs!Account_code, pos + 1))
                If suffix > LngMax Then
                    LngMax = suffix
                End If
            End If

            rs.MoveNext
        Loop

        suffix = LngMax + 1
    End If

    rs.Close
    Set rs = Nothing

    ' À«‰Ì«: ‰ √ﬂœ ≈‰ «·ﬂÊœ «··Ì Â‰Ê·œÂ „‘ „ÊÃÊœ ›Ì √Ì „ﬂ«‰ ›Ì «·ÃœÊ·
    Do
        NewCode = StrParentAccountCode & "a" & CStr(suffix)

        If Not AccountCodeExists(NewCode) Then
            Exit Do
        End If

        suffix = suffix + 1
    Loop

    GetNewAcountCode = NewCode
End Function


Private Function GetNewAcountCode2(StrParentAccountCode As String) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long, j As Long
    Dim LngMax As Long
    Dim pos As Integer

    StrSQL = "SELECT Account_Code "
    StrSQL = StrSQL + " From ACCOUNTS  Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccountCode & "'))"
    StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        GetNewAcountCode2 = StrParentAccountCode & "a" & 1
        Exit Function
    Else
        pos = find_a_pos(rs("Account_Code").value)

        LngMax = mId(rs("Account_Code").value, pos + 1, Len(rs("Account_Code").value) - pos)

        For i = 0 To rs.RecordCount - 1
            pos = find_a_pos(rs("Account_Code").value)

            If mId(rs("Account_Code").value, pos + 1, Len(rs("Account_Code").value) - pos) > LngMax Then
                LngMax = mId(rs("Account_Code").value, pos + 1, Len(rs("Account_Code").value) - pos)
            End If
     
            rs.MoveNext
        Next i

        GetNewAcountCode2 = StrParentAccountCode & "a" & (LngMax + 1)
    End If

End Function

Private Function GetNewAcountCodeold(StrParentAccountCode As String) As String
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Long, j As Long
    Dim LngMax As Long

    StrSQL = "SELECT Right(Account_Code,1)as no1,Right(Account_Code,2)as no2,Right(Account_Code,3)as no3 "
    StrSQL = StrSQL + ",Right(Account_Code,4)as no4,Right(Account_Code,5)as no5,Right(Account_Code,6)as no6 "
    StrSQL = StrSQL + " From ACCOUNTS  Where ((ACCOUNTS.Parent_Account_Code='" & StrParentAccountCode & "'))"
    StrSQL = StrSQL + " ORDER by ACCOUNTS.Account_ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        GetNewAcountCodeold = StrParentAccountCode & "a" & 1
        Exit Function
    Else

        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.count - 1

                If LngMax < val(rs(j).value) Then
                    LngMax = val(rs.Fields(j).value)
                End If

            Next j

            rs.MoveNext
        Next i

        GetNewAcountCodeold = StrParentAccountCode & "a" & (LngMax + 1)
    End If

End Function

Public Function GetNewDEV_Serial(RecordDate As Date) As String
    Dim StrSQL As String
    Dim RsSerial As ADODB.Recordset
    Dim LngSerialCount As Long
VBA.Calendar = vbCalGreg
    StrSQL = "Select Distinct Double_Entry_Vouchers_ID From " & " DOUBLE_ENTREY_VOUCHERS"
    StrSQL = StrSQL & " Where RecordDate=" & SQLDate(RecordDate, True) & ""
    Set RsSerial = New ADODB.Recordset
    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsSerial.BOF Or RsSerial.EOF Then
        LngSerialCount = 1
    Else
        LngSerialCount = RsSerial.RecordCount + 1
    End If

    GetNewDEV_Serial = year(RecordDate) & IIf(Len(Month(RecordDate)) = 1, "0" & Month(RecordDate), Month(RecordDate)) & IIf(Len(day(RecordDate)) = 1, "0" & day(RecordDate), day(RecordDate)) & "0" & LngSerialCount
End Function

Public Function DeleteAccount(StrAccountCode As String, _
                              Optional ChekDEV As Boolean = False) As Boolean
    Dim StrSQL  As String
    Dim Msg As String

    On Error GoTo hErr
    Dim sql As String

    Dim rs As New ADODB.Recordset

    sql = "select * from DOUBLE_ENTREY_VOUCHERS1 where Account_Code='" & StrAccountCode & "'"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GoTo hErr
    End If

    rs.Close

    If ChekDEV = True Then
        sql = "select * from DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            GoTo hErr
        End If
    End If

    StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    DeleteAccount = True
    Exit Function
hErr:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«»."
Else
Msg = "Can't Delete this Account"
End If
    MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DeleteAccount = False
End Function


Public Function CheckDeleteAccount(StrAccountCode As String, _
                              Optional ChekDEV As Boolean = False) As Boolean
    Dim StrSQL  As String
    Dim Msg As String

    On Error GoTo hErr
    Dim sql As String

    Dim rs As New ADODB.Recordset

    sql = "select * from DOUBLE_ENTREY_VOUCHERS1 where Account_Code='" & StrAccountCode & "'"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        GoTo hErr
    End If

    rs.Close

    If ChekDEV = True Then
        sql = "select * from DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            GoTo hErr
        End If
    End If

    'StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
    'Cn.Execute StrSQL, , adExecuteNoRecords
    CheckDeleteAccount = True
    Exit Function
hErr:
'If SystemOptions.UserInterface = ArabicInterface Then
'    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·Õ”«»."
'Else
'Msg = "Can't Delete this Account"
'End If
'    MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    CheckDeleteAccount = False
End Function


