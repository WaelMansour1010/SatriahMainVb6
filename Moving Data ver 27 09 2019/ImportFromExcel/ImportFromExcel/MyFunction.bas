Attribute VB_Name = "MyFunction"
'
'Public Sub GridFromToExecl(ByRef mGrid As Object, _
'                           frm As Form, _
'                           Optional ByVal mTable As String = "", _
'                           Optional ByVal ExtraFieldsName As String = "", Optional ByVal ExtraFieldsType As String = "", Optional ByVal ExtraFieldsTitle As String = "", Optional ByVal ExtraFieldsColComboList As String = "")
'    On Error Resume Next
'
'
'    Dim myform As New frm
'    Set myform = New frm
'
'  '  myform.mTableName = mTable
'    myform.mGrid.Rows = 1
'    Set myform.SenderObject = frm
'    Set myform.SenderGrid = mGrid
'    'myform.mGrid.Cols = 1
'    myform.mGrid.Rows = mGrid.Rows
'    '    myform.mGrid.Cols = mGrid.Cols
'    '    myform.mGrid2.Cols = mGrid.Cols
'    '    myform.mGrid3.Cols = mGrid.Cols
'
'    myform.MainFormName = frm.Name
'    Dim colKey
'    Dim ColDataType
'    Dim TextMatrix
'    Dim ColComboList
'    '
'    If ExtraFieldsName <> "" Then
'        colKey = Split(ExtraFieldsName, ",")
'        ColDataType = Split(ExtraFieldsType, ",")
'        TextMatrix = Split(ExtraFieldsTitle, ",")
'        ColComboList = Split(ExtraFieldsColComboList, ",")
'        For i = 1 To UBound(colKey)
'            AddToGrid myform.mGrid, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mGrid2, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mGrid3, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'            AddToGrid myform.mtmpGrd, colKey(i), TextMatrix(i), ColDataType(i), ColComboList(i)
'        Next
'    End If
'    For i = myform.mGrid.Cols - 1 To 1 Step -1
'        If myform.mGrid.ColHidden(i) Then
'            '            myform.mGrid.ColPosition(i) = myform.mGrid.Cols - 1
'            '            myform.mGrid2.ColPosition(i) = myform.mGrid2.Cols - 1
'            '            myform.mGrid3.ColPosition(i) = myform.mGrid3.Cols - 1
'            '            myform.mtmpGrd.ColPosition(i) = myform.mtmpGrd.Cols - 1
'        End If
'    Next
'
'
'    '    myform.mGrid2.Cols = myform.mGrid.Cols
'    '    myform.mGrid3.Cols = myform.mGrid.Cols
'
'    For j = 0 To mGrid.Cols - 1
'        Screen.MousePointer = vbHourglass
'        'If Not myForm.mGrid.ColHidden(j) Then
'        myform.mGrid.colKey(j) = mGrid.colKey(j)
'        myform.mGrid.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid.ColHidden(j) = mGrid.ColHidden(j)
'
'        myform.mGrid2.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'        myform.mGrid2.colKey(j) = mGrid.colKey(j)
'        myform.mGrid2.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid2.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid2.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid2.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid2.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid2.ColHidden(j) = mGrid.ColHidden(j)
'
'        myform.mGrid3.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'        myform.mGrid3.colKey(j) = mGrid.colKey(j)
'        myform.mGrid3.FixedAlignment(j) = mGrid.FixedAlignment(j)
'        myform.mGrid3.ColAlignment(j) = mGrid.ColAlignment(j)
'        myform.mGrid3.ColComboList(j) = mGrid.ColComboList(j)
'        myform.mGrid3.ColDataType(j) = mGrid.ColDataType(j)
'        myform.mGrid3.ColFormat(j) = mGrid.ColFormat(j)
'        myform.mGrid3.ColHidden(j) = mGrid.ColHidden(j)
'
'        For i = 0 To mGrid.Rows - 1
'
'            If InStr(1, mGrid.ColComboList(j), "#") And i <> 0 Then
'                myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'            Else
'                myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'            End If
'            DoEvents
'        Next
'
'        'End If
'    Next
'    Screen.MousePointer = vbNormal
'    'GridSerial myForm.mGrid
'    myform.Show 1
'    If myform.Tag = "OK" Then
'        Screen.MousePointer = vbHourglass
'        mGrid.Rows = 1
'        mGrid.Rows = myform.mGrid.Rows
'        For j = 1 To myform.mGrid.Cols - 1
'            For i = 1 To myform.mGrid.Rows - 1
'                mGrid.TextMatrix(i, j) = myform.mGrid.TextMatrix(i, j)
'                DoEvents
'            Next
'        Next
'    End If
'  '  GridSerial mGrid
'    Screen.MousePointer = vbNormal
'    Unload myform
'
'    '  myform.mGrid.Rows = 1
'    '    myform.mGrid.Cols = 1
'    '    myform.mGrid.Rows = mGrid.Rows
'    '    myform.mGrid.Cols = mGrid.Cols
'    '    myform.mGrid2.Cols = mGrid.Cols
'    '    myform.mGrid3.Cols = mGrid.Cols
'    '    myform.MainFormName = frm.Name
'    '    myform.mTableName = mTable
'    '
'    '    If mTable <> "" Then
'    '        myform.cmdSave.Visible = True
'    '        myform.SSTTab0.TabsPerRow = 3
'    '        myform.SSTTab0.Tab = 0
'    '    Else
'    '        myform.SSTTab0.TabsPerRow = 1
'    '        myform.SSTTab0.TabVisible(0) = True
'    '        myform.SSTTab0.TabVisible(1) = False
'    '        myform.SSTTab0.TabVisible(2) = False
'    '        myform.cmdSave.Visible = False
'    '    End If
'    '    myform.mGrid2.Cols = myform.mGrid.Cols
'    '    myform.mGrid3.Cols = myform.mGrid.Cols
'    '
'    '    For j = 0 To mGrid.Cols - 1
'    '        Screen.MousePointer = vbHourglass
'    '        'If Not myForm.mGrid.ColHidden(j) Then
'    '        myform.mGrid.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid.ColHidden(j) = mGrid.ColHidden(j)
'    '
'    '        myform.mGrid2.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'    '        myform.mGrid2.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid2.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid2.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid2.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid2.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid2.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid2.ColHidden(j) = mGrid.ColHidden(j)
'    '
'    '        myform.mGrid3.TextMatrix(0, j) = mGrid.TextMatrix(0, j)
'    '        myform.mGrid3.colKey(j) = mGrid.colKey(j)
'    '        myform.mGrid3.FixedAlignment(j) = mGrid.FixedAlignment(j)
'    '        myform.mGrid3.ColAlignment(j) = mGrid.ColAlignment(j)
'    '        myform.mGrid3.ColComboList(j) = mGrid.ColComboList(j)
'    '        myform.mGrid3.ColDataType(j) = mGrid.ColDataType(j)
'    '        myform.mGrid3.ColFormat(j) = mGrid.ColFormat(j)
'    '        myform.mGrid3.ColHidden(j) = mGrid.ColHidden(j)
'    '        For i = 0 To mGrid.Rows - 1
'    '            myform.mGrid.TextMatrix(i, j) = mGrid.TextMatrix(i, j)
'    '            DoEvents
'    '        Next
'    '        'End If
'    '    Next
'    '    Screen.MousePointer = vbNormal
'    '    'GridSerial myForm.mGrid
'    '    myform.Show 1
'    '    If myform.Tag = "OK" Then
'    '        Screen.MousePointer = vbHourglass
'    '        mGrid.Rows = 1
'    '        mGrid.Rows = myform.mGrid.Rows
'    '        For j = 1 To myform.mGrid.Cols - 1
'    '            For i = 1 To myform.mGrid.Rows - 1
'    '                mGrid.TextMatrix(i, j) = myform.mGrid.TextMatrix(i, j)
'    '                DoEvents
'    '            Next
'    '        Next
'    '    End If
'    '    GridSerial mGrid
'    '    Screen.MousePointer = vbNormal
'    '    Unload myform
'End Sub
'
'



'
''emp_contract_type  ContractID
'MaritalStatus
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
'     Dcbsex.AddItem "–ﬂ—"
'      Dcbsex.AddItem "√‰ÀÏ"
'       Dcbsex.AddItem "Male"
'     Dcbsex.AddItem "Female"
'       DcbMatrial.AddItem "Single"
'     DcbMatrial.AddItem "Married"
'
'JobTypeID
'JobTypeID,JobTypeName From TblEmpJobsTypes
'
'pasplace
'select  id,name  from jopstatus
'select  id,name  from Nationality
' select  id,name  from dean
'
'
'SELECT DISTINCT pasplace, pasplace AS pasplaceName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (pasplace IS NULL))"
'
'
'
'sql = "SELECT DISTINCT BankCode, BankCode AS BankCodeName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (BankCode IS NULL)) "
'
'sql = "SELECT DISTINCT BanckName, BanckName AS BanckNameName"
'sql = sql & " From dbo.TblEmployee"
'sql = sql & " WHERE     (NOT (BanckName IS NULL)) "

Public Function GetGridFileName(ByVal G As Object, Optional MainFormName As String = "") As String
    Dim GlobalGridName As String
    Dim IndexS As String
    Dim MainContainerName As String

    On Error Resume Next
    IndexS = G.Index

    MainContainerName = GetMainForm(G.Container)
    GlobalGridName = MainContainerName & "\" & G.Name & IndexS & MainFormName
    GlobalGridName = "Import"
    GetGridFileName = App.Path & GlobalGridName & ".xls"

End Function

Private Function ToHex(ByRef pstrMessage As String) As String

    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrHex As String

    llngMaxIndex = LenB(pstrMessage)

    For llngIndex = 1 To llngMaxIndex
        lstrHex = lstrHex & Right("0" & Hex(AscB(MidB(pstrMessage, llngIndex, 1))), 2)
    Next

    ToHex = lstrHex

End Function

Private Function FromHex(ByRef pstrHex As String) As String

    Dim llngMaxIndex As Long
    Dim llngIndex As Long
    Dim lstrMessage As String

    llngMaxIndex = Len(pstrHex)

    For llngIndex = 1 To llngMaxIndex Step 2
        lstrMessage = lstrMessage & ChrB("&h" & Mid(pstrHex, llngIndex, 2))
    Next

    FromHex = lstrMessage

End Function

Private Function Translate(ByRef pstrMessage As String, ByVal Key As String) As String

    Dim llngIndex As Long
    Dim llngMessageLength As Long
    Dim llngKeyLength As Long
    Dim lstrKey As String
    Dim llngKeyIndex As Long
    Dim lbytMessageByte As Byte
    Dim lbytKeyByte As Byte
    Dim llngMessageIndex As Long
    Dim lstrTranslation As String

    lstrKey = ToHex(Key)
    llngKeyLength = Len(lstrKey) \ 2

    If llngKeyLength = 0 Then Exit Function

    llngMessageLength = Len(pstrMessage) \ 2

    For llngIndex = 1 To llngMessageLength

        llngKeyIndex = (((llngIndex - 1) Mod llngKeyLength) * 2) + 1
        llngMessageIndex = ((llngIndex - 1) * 2) + 1

        lbytMessageByte = Int("&h" & Mid(pstrMessage, llngMessageIndex, 2))
        lbytKeyByte = Int("&h" & Mid(lstrKey, llngKeyIndex, 2))

        lstrTranslation = lstrTranslation & ToHex(ChrB(lbytMessageByte Xor lbytKeyByte))

    Next

    Translate = lstrTranslation

End Function

Public Function HexDecrypt(ByVal pstrMessage As String, ByVal Key As String) As String
    pstrMessage = Translate(pstrMessage, Key)
    pstrMessage = FromHex(pstrMessage)
    HexDecrypt = pstrMessage
End Function
Public Function HexEncrypt(ByVal pstrMessage As String, ByVal Key As String) As String
    pstrMessage = ToHex(pstrMessage)
    pstrMessage = Translate(pstrMessage, Key)
    HexEncrypt = pstrMessage
End Function

Public Function GetMainForm(ByVal Obj) As String
    Dim n As String
    On Error Resume Next
    n = Obj.Container.Name

    If n = "" Then
        GetMainForm = Obj.Name
    Else
        GetMainForm = GetMainForm(Obj.Container)
    End If
End Function
Public Sub ToExcel(ByRef mGrid As Object, _
                   frm As Form, _
                   Optional MainFormName As String = "")
    On Error GoTo EH

    Screen.MousePointer = vbHourglass
    For i = 1 To mGrid.Cols - 1

        If Not mGrid.ColHidden(i) Then
            If mGrid.ColDataType(i) = 0 Then
                If mGrid.ColComboList(i) <> "" Then
                    mGrid.ColDataType(i) = flexDTSingle
                Else
                    mGrid.ColDataType(i) = flexDTString
                End If

            End If




        End If


    Next
    ExportToExcel frm, mGrid, , , MainFormName

    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    MsgBox MyErrorHandler(Err.Number)

End Sub

Public Function MyErrorHandler(ErrNo As Long) As String
    mMsg = ""
    Select Case ErrNo
    
    Case 0
        MyErrorHandler = ""
        Exit Function
 
    Case -2147217864

        If ArabicInterface Then
            mMsg = " „ ≈Ã—«¡  ⁄œÌ·«  ⁄·Ï Â–Â «·‘«‘Â „‰ ÃÂ«“ ¬Œ—- „‰ ›÷·ﬂ «⁄œ  Õ„Ì· «·Õ—ﬂÂ À„ Õ«Ê· „—Â «Œ—Ï" & " - Optimistic concurrency erorr "
        Else
            mMsg = "This Form is editing from another computer- realod and Try again " & " - Optimistic concurrency erorr "
        End If

    Case -2147467259
        If ArabicInterface Then
            mMsg = "«·ÃÂ«“ «·Œ«œ„ «·—∆Ì”Ì „€·ﬁ √Ê €Ì— „ÊÃÊœ ⁄·Ï Â–Â «·‘»ﬂ…" & " - " & ErrNo
        Else
            mMsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case -2147352567
        'If ArabicInterface Then
        '    mMsg = "ÌÃ»  Œ’Ì’ «·ÿ«»⁄«  „‰ ≈œ«—… «·‰Ÿ«„" & " - " & ErrNo
        'Else
        '    mMsg = "Select Correct Report Printer Device" & " - " & ErrNo
        'End If
    Case 3155, 3022, -2147217873, -2147217900    ' insert fail
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· ° Â–Â «·»Ì«‰«   „  ”ÃÌ·Â« „‰ ﬁ»·" & " - " & ErrNo
        Else
            mMsg = "You Can not Add this Record , May be there is Dublicated values" & " - " & ErrNo
        End If
    Case 3200    ' Change Or Delete Failed
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «·€«¡ √Ê  ⁄œÌ· Â–« «·”Ã·  »”»» ÊÃÊœ »Ì«‰«  √Œ—Ï „— »ÿ… »Â ÊÌÃ» «·€«¡Â« √Ê·«" & " - " & ErrNo
        Else
            mMsg = "You Can not Delete Or Modify this Record , Because There Is Some Data Depends On It " & " - " & ErrNo
        End If
    Case 3157, 3046, 3202, 3218    ' Update Fail
        If ArabicInterface Then
            mMsg = " Â‰«ﬂ ›‘· ›Ï  Œ“Ì‰ «· ⁄œÌ·«  ° ﬁœ ÌﬂÊ‰ «·”Ã· „ﬁ›· »Ê«”ÿ… „” Œœ„ ¬Œ—° Õ«Ê· „—… √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Update Failed , May be The record is locked by another User , Try Again " & " - " & ErrNo
        End If
    Case 3186, 3187, 3188
        If ArabicInterface Then
            mMsg = "”Ã· „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            mMsg = "Current Record locked by Another user" & " - " & ErrNo
        End If
    Case 3167
        If ArabicInterface Then
            mMsg = " „ «·€«¡ Â–« «·”Ã· »«·›⁄· " & " - " & ErrNo
        Else
            mMsg = "Record Already Deleted" & " - " & ErrNo
        End If
    Case 3314
        If ArabicInterface Then
            mMsg = "„‰ ›÷·ﬂ √ﬂ„· «·»Ì«‰«  ﬁ»· «· Œ“Ì‰" & " - " & ErrNo
        Else
            mMsg = "Please Complete the data before saving" & " - " & ErrNo
        End If
    Case 3262, 3211, 3212    ' Locked by another user and wait
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰ ≈€·«ﬁ «·„·› »”»» ÊÃÊœ „” Œœ„ ¬Œ— ÌﬁÊ„ »≈” Œœ«„Â √Ê ﬁ«„ »≈€·«ﬁÂ" & " - " & ErrNo
        Else
            mMsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case 3197    ' Couldn't repaire this files
        If ArabicInterface Then
            mMsg = "√ﬂÀ— „‰ „” Œœ„ Õ«Ê·Ê«  €ÌÌ— ‰›” «·»Ì«‰«  ›Ï ‰›” «·Êﬁ " & " - " & ErrNo
        Else
            mMsg = "Another Users are attempting to change the same data at the same time" & " - " & ErrNo
        End If
    Case 3056    ' Couldn't repaire this files
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰  ’·ÌÕ «·„·›«  «·„” Œœ„…" & " - " & ErrNo
        Else
            mMsg = "Couldn't repaire this files" & " - " & ErrNo
        End If
    Case 3014, 3037    ' Can't open any more files
        If ArabicInterface Then
            mMsg = "·« Ì„ﬂ‰ › Õ „·›«  √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Can't open any more files" & " - " & ErrNo
        End If
    Case 3356, 3260, 3261, 3189, 3008, 3164, 3006    ' Table or Database Locked
        If ArabicInterface Then
            mMsg = "«·„·› „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            mMsg = "The File is Locked by Another User" & " - " & ErrNo
        End If
    Case 3201    ' Add Or Edit Fail
        If ArabicInterface Then
            mMsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· √Ê «· ⁄œÌ· ›ÌÂ ° ·√‰Â „— »ÿ »„·› ·„ Ì „ «·≈÷«›… √Ê «· ⁄œÌ· ›ÌÂ Õ Ï «·¬‰" & " - " & ErrNo
        Else
            mMsg = "You Can not Add this Record or Change it , Because it's Linked to a File that has not been Added or Changed till Now" & " - " & ErrNo
        End If
    Case -2147217887
        If ArabicInterface Then
            mMsg = "Œÿ√ €Ì— „⁄—Ê› ° Õ«Ê·  ‰›Ì– ‰›” «·⁄„·Ì… „—… √Œ—Ï" & " - " & ErrNo
        Else
            mMsg = "Undefined Error , Try again : " & ErrNo
        End If
    Case 3704
        
        On Error Resume Next
        db.Close
        Exit Function
    Case -1000000001
       
        MyErrorHandler = ""
        Exit Function
    End Select
    '*************************
    If Err.Number = vbObjectError + 1000 Then
        

        mMsg = mMsg & vbNewLine & Err.Description
    Else
        mMsg = mMsg & vbNewLine & Err.Description & " : " & Err.Number
    End If
    '*************************
    If ErrNo <> -2147217864 Then
        If db.Errors.Count > 0 Then
            ss = ""
            Dim adoErr As ADODB.Error
            j = 1
            On Error GoTo EEE
            For Each adoErr In db.Errors
                If adoErr.Number <> 0 Then
                    If j = 1 Then ss = vbNewLine & "-------SQL Errors-------"
                    ss = ss & vbNewLine & "Error (" & j & ")=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
                    j = j + 1
                End If
            Next adoErr
EEE:
            ' for this rand error Not enough storage is available to process this command.
            If Err.Number = 48 Then
                Set adoErr = db.Errors(0)
                ss = ss & vbNewLine & "Error (48)=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
            End If
            On Error GoTo 0
            mMsg = mMsg & vbNewLine & ss
        End If
    End If
    '*************************
    'If Trim(mMsg) <> "()(0)" Then MyErrorHandler = mMsg Else MyErrorHandler = ""
    MyErrorHandler = mMsg & ":" & Erl
    IsAboutError = True

End Function




Public Sub ExportToExcel(ByRef frm As Form, _
                         ByRef G As Object, _
                         Optional Caption As String = "", _
                         Optional ByRef ObjExecel, Optional MainFormName As String = "", Optional mRtl As Integer = 0)

'***********************Khalid
'    If Not isDebugMode Then
'        If IsDisabledExcelShortCut Then
'            Exit Sub
'        End If
'    End If
    '****************************

    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    '--- open new Excel File In memory ---
    Set ExcelSheet = CreateObject("excel.application")
    If Not IsObject(ExcelSheet) Then Exit Sub

    '--- Add new WorkBook ---
    '--- ByDefault contain 3 Woorksheet ---
    '    ExcelSheet.Workbooks.Add
    '
    'For i = 0 To G.Cols - 1
    'G.ColDataType(i) = flexDTStringC
    'Next
    '    '==== FIRST WORKSHEET ===================================
    '    '--- activate the second WorkSheet ---
    'FFF = "c:\~Temp" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
    FFF = GetGridFileName(frm.Grd, MainFormName)
    frm.Grd.SaveGrid FFF, flexFileExcel, flexXLSaveFixedRows Or flexXLSaveFixedCols
  '  ExcelSheet.Workbooks.Open FFF
        Set ExcelObj = CreateObject("Excel.Application")
    '
  '  Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open FFF
    '    ExcelSheet.ActiveWorkbook.Sheets("Sheet2").Delete
    '    ExcelSheet.ActiveWorkbook.Sheets("Sheet3").Delete
    '    ExcelSheet.ActiveWorkbook.Sheets("Sheet1").Activate
    '    ExcelSheet.ActiveWorkbook.Sheets("Sheet1").Name = Trim(mId(frm.caption, InStrRev(frm.caption, ":") + 1))

    '    For i = 0 To G.Rows - 1
    '    jj = 1
    '        For j = 0 To G.Cols - 1
    '            If Not G.ColHidden(j) Then
    '                ExcelSheet.Cells(i + 1, jj).value = G.TextMatrix(i, j)
    '                ExcelSheet.Columns(jj).Font.Bold = True
    '                ExcelSheet.Columns(jj).Font.Size = 12
    '                ExcelSheet.Columns(jj).Borders.Color = RGB(200, 200, 200)
    '                ExcelSheet.Columns(jj).Interior.Color = RGB(255, 255, 215)
    '                ExcelSheet.Cells(1, jj).Interior.Color = RGB(153, 204, 255)
    '                ExcelSheet.Cells(1, jj).Font.Color = RGB(255, 255, 255)
    '                ExcelSheet.Cells(1, jj).Font.Size = 14
    '                jj = jj + 1
    '            End If
    '        Next
    '    Next
    'ExcelSheet.Cells( i, ).Interior.Color

    For j = 0 To G.Cols - 1
        If Not G.ColHidden(j) Then
            ExcelSheet.Columns(j + 1).Font.Bold = True
            ExcelSheet.Columns(j + 1).Font.Size = 12
            ExcelSheet.Columns(j + 1).Borders.Color = RGB(200, 200, 200)
            ExcelSheet.Columns(j + 1).Interior.Color = RGB(255, 255, 215)
            ExcelSheet.cells(1, j + 1).Interior.Color = RGB(153, 204, 255)
            ExcelSheet.cells(1, j + 1).Font.Color = RGB(255, 255, 255)
            ExcelSheet.cells(1, j + 1).Font.Size = 14
        End If
    Next

    For j = 1 To G.Cols - 1
        If G.ColHidden(j) Or G.ColWidth(j) = 0 Then
            ExcelSheet.Columns(j + 1).Clear
        End If
    Next
    For j = G.Cols - 1 To 1 Step -1
        If G.ColHidden(j) Or G.ColWidth(j) = 0 Then
            ExcelSheet.Columns(j + 1).Delete
        End If
    Next

    On Error Resume Next
    For j = G.Cols - 1 To 1 Step -1
        If G.ColComboList(j) <> "" Then
            OpValue = Split(G.ColComboList(j), ";")
            OpCaption = Split(G.ColComboList(j), "|")
            If UBound(OpValue) > 0 And UBound(OpCaption) > 0 Then
                OpCaptionss = ""

                zz = UBound(OpCaption)
                ReDim OpCaptionsss(zz) As String
                For i = 0 To UBound(OpCaption)
                    OpCaptions = Split(OpCaption(i), ";")
                    If i > 0 Then OpCaptionss = OpCaptionss & ";"
                    OpCaptionss = OpCaptionss & OpCaptions(1)

                Next
                'OpValue = Split(tRs!OptionsValue, ":")
                'OpCaption = Split(IIf(ArabicInterface, tRs!OptionsArabic, tRs!OptionsEnglish), ":")
                'OpCaption = Replace(IIf(ArabicInterface, tRs!OptionsArabic, tRs!OptionsEnglish), ":", ";")
                '5-4-2015
                ' „  ‰ﬁÌ’ ﬁÌ„… «· J
                OpCaptionsss = Split(OpCaptionss, ";")
                hid = 0
                For jj = j - 1 To 1 Step -1
                    hid = hid + IIf(G.ColHidden(jj), 1, 0)
                Next
                'ExcelSheet.Columns(j + 1 - hid).Validation.Delete
                'ExcelSheet.Columns(j + 1 - hid).Validation.Add xlValidateList, , , OpCaptionss
                Dim MyLS As String
                MyLS = GetListSeparator
                With ExcelSheet.Columns(j + 1 - hid).Validation
                    .Delete

                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                         Operator:=xlEqual, Formula1:=Join(OpCaptionsss, MyLS)
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With

            End If
        End If
    Next
    On Error GoTo EH

    '--- Change Width For All Columns Automatic ---
    ExcelSheet.Columns.AutoFit
    '    '========================================================
    If mRtl = 0 Then
        ExcelSheet.Worksheets(1).DisplayRightToLeft = ArabicInterface
    ElseIf mRtl = 1 Then
        ExcelSheet.Worksheets(1).DisplayRightToLeft = True
    Else
        ExcelSheet.Worksheets(1).DisplayRightToLeft = False
    End If
    '--- Show the Excel File ---
    ExcelSheet.Visible = True
    ObjExecel = FFF
    LastExportedExcelFile = FFF
    Screen.MousePointer = vbDefault
    Exit Sub

EH:
    Screen.MousePointer = vbDefault
    MsgBox "Â‰«ﬂ «Œÿ«¡ ° —»„« »”»» ⁄œ„ ÊÃÊœ «·«ﬂ”Ì· √Ê «·„·› „› ÊÕ „”»ﬁ«"

End Sub




Public Sub FromExcel(ByRef mGrid As Object, _
                     ByRef mtmpGrd As Object, _
                     frm As Form, _
                     Optional MainFormName As String = "", _
                     Optional ProgressBar As Object = Nothing, Optional ByVal XlsFileName As String = "", Optional ByVal MainTableName As String = "")


    ' If Not i Then Exit Sub
       Dim cProgress As ClsProgress
    '    Dim mtmpGrd As VSFlexGrid
    If XlsFileName = "" Then
        XlsFileName = GetGridFileName(mGrid, MainFormName)
    End If
    If FileExists(XlsFileName) Then

        mtmpGrd.FixedCols = 0
        mtmpGrd.FixedRows = 0

        mtmpGrd.LoadGrid XlsFileName, flexFileExcel

        mtmpGrd.BackColor = &HFFFFFF
        mtmpGrd.BackColorAlternate = &HE9E9E9
        mtmpGrd.BackColorBkg = &H8000000C
        mtmpGrd.BackColorFixed = &H8000000F
        mtmpGrd.BackColorFrozen = &HC0FFFF
        mtmpGrd.BackColorSel = &H8000000D
        mtmpGrd.ForeColor = &H80000008
        mtmpGrd.ForeColorFixed = &HFF0000
        mtmpGrd.ForeColorSel = &H8000000E
        mtmpGrd.GridColor = &H8000000F
        mtmpGrd.GridColorFixed = &H80000010
        mtmpGrd.FixedCols = 1
        mtmpGrd.FixedRows = 1
        '·«‰ Loaded ÌŒ ›Ì
        mtmpGrd.Cols = mtmpGrd.Cols + 1
        mtmpGrd.ColKey(mtmpGrd.Cols - 1) = "Loaded"
        mtmpGrd.ColHidden(mtmpGrd.Cols - 1) = True
        mtmpGrd.AutoSize 0, mtmpGrd.Cols - 1
    End If
    mGrid.Rows = 1
    mGrid.Rows = mtmpGrd.Rows

    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Min = 1
        ProgressBar.Max = IIf(mGrid.Rows > 2, mGrid.Rows - 1, 2)    ' mGrid.Rows - 1
        ProgressBar.Visible = True
        '********************************
    End If
        Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    



       
    
    For i = 1 To mtmpGrd.Rows - 1
        '********************************
        If Not ProgressBar Is Nothing Then
            ProgressBar.Value = i
            DoEvents
            ProgressBar.Refresh
        End If
        cProgress.StartProgress
       DoEvents
        '********************************
        jj = 0
        For j = 1 To mGrid.Cols - 1
            If j = 18 Then
                j = 18
            End If
            If Not mGrid.ColHidden(j) Then
                jj = jj + 1
                       If mGrid.ColKey(j) = "JobTypeID2Name" Then
                    j = j
                End If
                If InStr(1, mGrid.ColComboList(j), "#") Then
                    Hide = 0
                    For H = j - 1 To 1 Step -1
                        Hide = Hide + IIf(mGrid.ColHidden(H), 1, 0)
                    Next
                    mGrid.TextMatrix(i, j) = mtmpGrd.TextMatrix(i, j - Hide)
                    'Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                Else
                    mGrid.TextMatrix(i, j) = Replace(Trim(mtmpGrd.TextMatrix(i, jj)), "'", "")
                End If
                If Trim(mGrid.ColEditMask(j)) = "Date" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid
                End If
                'pValue = Split(G.ColComboList(j), ";")
            Else
                j = j
                If j = 34 Then
                j = j
                End If
                If Trim(mGrid.ColEditMask(j)) <> "" Then
                    GetFieldID mGrid.ColEditMask(j), i, j, mGrid, MainTableName
                End If
                If Trim(mGrid.ColComboList(j)) <> "" Then
                    GetIDCombo Trim(mGrid.ColComboList(j)), i, j, mGrid
                End If
            End If
            If Trim(Replace(Trim(mtmpGrd.TextMatrix(i, 1)), "'", "")) = "" Then
                mGrid.Rows = i + 1: Exit Sub
            End If
        Next
        ' DisplayOrderTotals
    Next
    '********************************
    If Not ProgressBar Is Nothing Then
        ProgressBar.Visible = False
    End If
           DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    MsgBox " „ «·«œ—«Ã"
    '********************************
    
End Sub

Private Sub GetIDCombo(ByVal mTableColID As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object)
Dim mTxt As String
mTxt = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Select Case mTableColID
Case "sexID"
    If mTxt = "Male" Or mTxt = "–ﬂ—" Then
        mTxt = 1
    Else
        mTxt = 2
    End If
Case "MaritalStatusID"
'    DcbMatrial.AddItem "√⁄“»"
'      DcbMatrial.AddItem "„ “ÊÃ"
    If mTxt = "√⁄“»" Or mTxt = "Single" Then
        mTxt = 0
    ElseIf mTxt = "„ “ÊÃ" Or UCase(mTxt) = "MARRIED" Then
        mTxt = 1
    ElseIf mTxt = "„ÿ·ﬁ/„ÿ·›…" Or UCase(mTxt) = "DIVORCED" Then
        mTxt = 2
    ElseIf mTxt = "«—„·/√—„·…" Or UCase(mTxt) = "WIDOWED" Then
        mTxt = 3
        
    End If
Case "Emp_Name1.Emp_Name2.Emp_Name3.Emp_Name4"
    mTxt = mGrid.TextMatrix(mRow, mCol - 4) + " " + mGrid.TextMatrix(mRow, mCol - 3) + " " + mGrid.TextMatrix(mRow, mCol - 2) + " " + mGrid.TextMatrix(mRow, mCol - 1)
Case ""
End Select
mGrid.TextMatrix(mRow, mCol) = mTxt
End Sub

Public Function ToHijriDate(ByVal GregorianDate As String) As String
    Dim HijriDate As String, DateFormat As String
    ' DateFormat = "long date"
    
    DateFormat = "dd-mm-yyyy"
    HijriDate = ConvertDate(GregorianDate, vbCalGreg, vbCalHijri, DateFormat)
    ToHijriDate = HijriDate
    
End Function
Private Function ConvertDate(ByRef StringIn As String, _
                             ByRef OldCalender As Integer, _
                             ByVal NewCalender As Integer, _
                             ByRef NewFormat As String) As String
                             If StringIn = "" Then Exit Function
On Error Resume Next
    Dim SavedCal As Integer
    Dim d As Date, s As String
    SavedCal = Calendar
    Calendar = OldCalender
    d = CDate(StringIn)
    Calendar = NewCalender
    s = CStr(d)
    ConvertDate = Format(s, NewFormat)
    Calendar = SavedCal
End Function

Public Function ToGregorianDate(ByVal HijriDate As String) As Date
    Dim GregorianDate As String, DateFormat As String
  If HijriDate = "" Then Exit Function
    DateFormat = "dd/mm/yyyy"
    
    GregorianDate = ConvertDate(HijriDate, vbCalHijri, vbCalGreg, DateFormat)
    If DateDiff("D", "01/01/1900", GregorianDate) < 0 Then
    GregorianDate = Date
    End If
    ToGregorianDate = GregorianDate
End Function

Public Function CheckDateIsHij(ByVal mDate As String) As Integer
    If Not IsDate(mDate) Then CheckDateIsHij = 3: Exit Function
    
    If Trim(mDate) = "" Then CheckDateIsHij = 3: Exit Function
    
    If Year(mDate) < 1800 Then
        CheckDateIsHij = 1
    Else
        CheckDateIsHij = 2
    End If
End Function
Private Sub GetFieldID(ByVal mTableColName As String, ByVal mRow As Long, ByVal mCol As Long, ByVal mGrid As Object, Optional ByVal MainTableName As String = "")
    Dim mTableName As String
    Dim mFieldIDName As String
    Dim mFieldName As String
    Dim xx As Variant
    Dim mValue As String
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummy2 As New ADODB.Recordset
    If mCol = 67 Then
        mCol = 67
    End If
    If mGrid.ColKey(mCol) = "NationlID" Then
        mCol = mCol
    End If
    Dim mValue2 As String
    If mGrid.ColKey(mCol) = "DeanID" Then
        mCol = mCol
    End If
    If mGrid.ColKey(mCol) = "DOBH" Then
        mCol = mCol
    End If
    If mTableColName = "Date" Then
        If CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 1 Then
            'If Trim(mGrid.TextMatrix(mRow, mCol - 1)) <> "" Then
                mGrid.TextMatrix(mRow, mCol) = Trim(mGrid.TextMatrix(mRow, mCol - 1))
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
            'Else
            'End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 2 Then
            If Trim(mGrid.TextMatrix(mRow, mCol - 1)) = "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            Else
                mGrid.TextMatrix(mRow, mCol) = ToHijriDate(Trim(mGrid.TextMatrix(mRow, mCol - 1)))
            End If
        ElseIf CheckDateIsHij(Trim(mGrid.TextMatrix(mRow, mCol - 1))) = 3 Then
            If mGrid.TextMatrix(mRow, mCol) <> "" Then
                mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(Trim(mGrid.TextMatrix(mRow, mCol)))
            End If
            'mGrid.TextMatrix(mRow, mCol - 1) = ToGregorianDate(mGrid.TextMatrix(mRow, mCol))
        Else
        
        End If
        Exit Sub
    End If
    xx = Split(mTableColName, ",")
    mTableName = xx(0)
    mFieldIDName = xx(1)
    mFieldName = xx(2)
    
    mValue = Trim(mGrid.TextMatrix(mRow, mCol - 1))
Dim sss As String
sss = ""
Dim mValue3 As String

mValue3 = mValue
If (Right(mValue, 1)) = "Â" Then
    sss = "…"
ElseIf (Right(mValue, 1)) = "…" Then
    sss = "Â"
End If
If sss <> "" Then
    mValue3 = Replace(mValue3, Right(mValue3, 1), s)
End If

    Dim s As String
    mValue2 = mValue
    Select Case mTableName
    Case "jopstatus"
        If UCase(mValue) = "ACTIVE" Then
            mValue2 = "⁄·Ï ﬁÊ… «·⁄„·"
            
        End If
    Case "dean"
      If UCase(mValue) = "ISLAM" Then
            mValue2 = "„”·„"
       ElseIf UCase(mValue) = "CHRISTIAN" Then
            mValue2 = "„”ÌÕÏ"
        End If
    Case "Nationality"
        If UCase(mValue) = "JORDAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "INDIA" Then
            mValue2 = "Â‰œ"
        ElseIf Trim(UCase(mValue)) = "" Then
            mValue2 = "”⁄ÊœÌ"
        ElseIf UCase(mValue) = "EGYPT" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "PAKISTAN" Then
            mValue2 = "»«ﬂ” «‰"
        ElseIf UCase(mValue) = "BANGLADESH" Then
            mValue2 = "»‰Ã·«œÌ‘"
        ElseIf UCase(mValue) = "SUDAN" Then
            mValue2 = "”Êœ«‰"
        ElseIf UCase(mValue) = "ETHIOPIA" Then
            mValue2 = "«ÀÌÊ»Ì«"
            
        ElseIf UCase(mValue) = "CAMEROON" Then
            mValue2 = "ﬂ«„Ì—Ê‰"
        ElseIf UCase(mValue) = "PALESTINE" Then
            mValue2 = "›·”ÿÌ‰"
        ElseIf UCase(mValue) = "SYRIA" Then
            mValue2 = "”Ê—Ì«"
        ElseIf UCase(mValue) = "JORDANIAN" Then
            mValue2 = "«—œ‰"
        ElseIf UCase(mValue) = "AMERICA" Then
            mValue2 = "«„—Ìﬂ«"
        ElseIf UCase(mValue) = "EGYPTIAN" Then
            mValue2 = "„’—"
        ElseIf UCase(mValue) = "KENYA" Then
            mValue2 = "ﬂÌ‰Ì«"
        ElseIf UCase(mValue) = "LEBANON" Then
            mValue2 = "·»‰«‰"
        ElseIf UCase(mValue) = "SIRLANKIAN" Then
            mValue2 = "”Ì—·«‰ﬂ"
        ElseIf UCase(mValue) = "YEMEN" Then
            mValue2 = "Ì„‰"
        ElseIf UCase(mValue) = "TUNIS" Then
            mValue2 = " Ê‰”"
        ElseIf UCase(mValue) = "MALAYSIA" Then
            mValue2 = "„«·Ì“Ì«"
         Else
            mValue2 = mValue
         
            
        End If
        If mValue = "" Then mValue2 = "”⁄ÊœÌ"
    Case Else
    End Select
    If mValue = "" Then
        Exit Sub
    End If
    s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & "e   "
    If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Then
        s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
    End If
    s = s & " from  " & mTableName
    s = s & " Where (" & mFieldName & " = '" & Trim(mValue2) & "' Or " & Trim(mFieldName) & "e = '" & Trim(mValue) & "')"
    s = s & " or (" & mFieldName & " = '" & Trim(mValue3) & "' Or " & Trim(mFieldName) & "e = '" & Trim(mValue3) & "')"
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    If rsDummy.EOF Then
        s = s & " Or ( " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & "e Like '%" & Trim(mValue) & "%')"
    End If
    rsDummy.Close
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    If UCase(mTableName) = "GROUPS" And rsDummy.EOF Then
        rsDummy.Close
             s = "Select " & mFieldName & " ," & mFieldIDName & " ," & Trim(mFieldName) & "e   "
        If UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Then
            s = s & " ,ParentID,FullCode,GroupCode,Code,LastGroup "
        End If
        Dim mValue4  As String
        mValue4 = Trim(mGrid.TextMatrix(mRow, mCol - 2))
        
        s = s & " from  " & mTableName
        s = s & " Where " & mFieldName & " Like '%" & Trim(mValue2) & "%' Or " & Trim(mFieldName) & "e Like '%" & Trim(mValue) & "%'"
        s = s & " Or Fullcode   Like '%" & Trim(mValue4) & "%' Or Code Like '%" & Trim(mValue4) & "%'"
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        If rsDummy.EOF Then
            mValue4 = mValue4
        End If
    End If
    
    If Not rsDummy.EOF Then
        mGrid.TextMatrix(mRow, mCol) = Val(rsDummy.Fields.Item(Trim(mFieldIDName)) & "")
        If mGrid.ColKey(mCol) = "ParentID" Then
            mGrid.TextMatrix(mRow, mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(mRow, mGrid.ColIndex("FullCode")))
            Dim mmm As String
            mmm = SearchInGrid(mGrid, mValue, "GroupName")
            If mmm <> "" Then
                'mGrid.TextMatrix(mRow, mGrid.ColIndex("GroupCode")) = GetNewGroupCode(Val(mGrid.TextMatrix(CLng(mmm), mGrid.ColIndex("NewId"))))
            End If
            mGrid.TextMatrix(mRow, mGrid.ColIndex("LastGroup")) = 0
        End If

    Else
       
        rsDummy.AddNew
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy(Trim(mFieldName) & "e") = mValue
        If mGrid.ColKey(mCol) = "ParentID" Then
            'rsDummy("ParentID") = mValue
            Dim mm As String
            mm = SearchInGrid(mGrid, mValue, "GroupName")
            If mm <> "" Then
                rsDummy("ParentID") = Val(mGrid.TextMatrix(CLng(mm), mCol))
                rsDummy("FullCode") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                rsDummy("Code") = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
            End If
            rsDummy("GroupCode") = GetNewGroupCode(Val(rsDummy("ParentID") & ""), mTableName)
            rsDummy("LastGroup") = 0
            If mm <> "" Then
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("Code")) = Trim(mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("FullCode")))
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("GroupCode3")) = rsDummy("GroupCode") & ""
                mGrid.TextMatrix(CLng(mm), mGrid.ColIndex("LastGroup")) = 0
            End If
        End If
        s = "Select Max(" & mFieldIDName & ")  as MaxID  from  " & mTableName
        
        rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
        Dim mMaxId As Long
        If Not rsDummy2.EOF Then
            mMaxId = Val(rsDummy2!MaxId & "") + 1
        Else
            mMaxId = 1
        End If
        If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
            rsDummy(Trim(mFieldIDName)) = mMaxId
        End If
        rsDummy(Trim(mFieldName)) = mValue
        rsDummy.Update
       ' mGrid.TextMatrix(mRow, mGrid.ColIndex("NewId")) = mMaxId
        mGrid.TextMatrix(mRow, mCol) = rsDummy(Trim(mFieldIDName) & "")
    End If

End Sub

Private Function SearchInGrid(ByVal mGrd As Object, ByVal mTxt As String, ByVal mFldName As String) As String
Dim i As Long
For i = 1 To mGrd.Rows - 1
    If Trim(mGrd.TextMatrix(i, mGrd.ColIndex(mFldName))) = mTxt Then
        SearchInGrid = i
        Exit Function
    End If
Next
SearchInGrid = ""
End Function
Function FileExists(FileName) As Boolean
    On Error GoTo CheckError        ' Turn on error trapping so error handler                            ' responds if any error is detected.
    FileExists = (Dir(FileName) <> "")
    Exit Function            ' Avoid executing error handler                             ' if no error occurs.

CheckError:        ' Branch here if error occurs.    ' Define constants to represent Visual Basic error code.
    FileExists = False
    Resume Next
End Function



Public Function ConnectionFirst() As Boolean

On Error GoTo ErrTrap
'«” ›”«—
'ServerDb = TxtServerDataBaseName.Text
'wael
'ServerDb = DestinationServer
' POSDb = TxtServerDataBaseName.Text




     Set Cn = New ADODB.Connection
    With Cn
        .CommandTimeout = 60
        .CursorLocation = adUseClient
        .ConnectionTimeout = 15
       If SysSQLServerType = 1 Then
        .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
        "Persist Security Info=False;Initial Catalog=" & ServerDb & _
        ";Data Source=" & SysSQLServerName & ";Port=1433"
        
        ElseIf SysSQLServerType = 2 Then
 
     
                 If SysSQLServerTypeTechnical = "0" Then
                 .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & _
                    "Persist Security Info=False;Initial Catalog=" & ServerDb & _
                    ";Data Source=" & SysSQLServerName & ";Port=1433"
                    '";Data Source=" & ServerDb & ";Port=1433"
                    
                  Else
                     .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SysSQLServerUserId & ";Initial Catalog=" & ServerDb & ";Data Source=" & SysSQLServerName 'SysSQLServerName
                End If
          End If

.Open
End With
ConnectionFirst = True
Exit Function
ErrTrap:
ConnectionFirst = False
End Function





 
 
 Public Sub saveGridExcel(ByVal Sqlstmt As String, ByRef tGrd As Object, ByVal ChekPoint As String, ByVal Index As String, mTableName As String, ParamArray FieldValue())
    On Error GoTo Err
    Dim tRs As New ADODB.Recordset
    Dim s As String
    Dim BolFrmLoaded As Boolean
    Dim mIndex As Long
     Dim cProgress As ClsProgress
    Dim mLastIndex As String
    If Index <> "" Then
        If Mid(Index, 1, 5) = "Index" Then
            mLastIndex = Mid(Index, 6)
            Index = "Id"
        End If
    End If
    Dim mMaxId As Long
    Dim rsDummy As New ADODB.Recordset
        s = "SELECT Max(" & Index & ") MaxID  FROM " & mTableName & " AS te "
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        If Not rsDummy.EOF Then
            mMaxId = Val(rsDummy!MaxId & "")
        End If
    
    
    tRs.Open Sqlstmt, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    ' *******************************************
    Dim II As Integer, i As Integer
    If Val(mLastIndex) = 0 Then
        II = 0
    Else
        II = Val(mLastIndex)
    End If
          Set cProgress = New ClsProgress
       cProgress.ProgressType = Waiting
    

    For i = tGrd.FixedRows To tGrd.Rows - 1
    



       cProgress.StartProgress
       DoEvents
      
 

        If BolFrmLoaded = True Then
           ' cProgress.StopProgess
           ' Set cProgress = Nothing
        End If
        
        If ChekPoint <> "" Then
            If Trim(tGrd.TextMatrix(i, tGrd.ColIndex(ChekPoint))) = "" Then GoTo NextStep
        End If
        Set rsDummy = New ADODB.Recordset
        s = "SELECT * FROM " & mTableName & " AS te WHERE te." & ChekPoint & " = N'" & Trim(tGrd.TextMatrix(i, tGrd.ColIndex(ChekPoint))) & "'"
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
         If mTableName <> "TblItemsUnits" Then
        If Not rsDummy.EOF Then GoTo NextStep
        End If
        '**********************
        tRs.AddNew
        II = II + 1
        mMaxId = mMaxId + 1
        If Index <> "" Then
        
            If mTableName = "tblItems" Then
                tRs!ItemId = mMaxId
                tRs!HaveSerial = 0
                tRs!HaveGuarantee = 0
                tRs!DealerPrice = 0
                tRs!GuaranteeValue = 0
                tRs!GuaranteeType = 0
                tRs!IsArchive = 0
                tRs!ItemType = 0
                tRs!AssbliedItem = 0
                tRs!RelatedItem = 0
                tRs!ItemCase = 1
                tRs!AssbliedItem = 0
                
                tGrd.TextMatrix(i, tGrd.ColIndex("ItemID")) = mMaxId
            ElseIf UCase(mTableName) = "GROUPS" Or UCase(mTableName) = "GROUPSCUSTOMERS" Then
                tRs!GroupCode = GetNewGroupCode(Val(tGrd.TextMatrix(i, tGrd.ColIndex("ParentID"))), mTableName)
                If UCase(mTableName) <> "GROUPSCUSTOMERS" Then
                    tRs!GroupId = Val(mMaxId)
                End If
            ElseIf UCase(mTableName) = "TBLCUSTEMERS" Then
                tRs!CusID = mMaxId
                
               ' tRs!Type = IIf(Option2.v, 1, 2)
                tRs!CreditlimitCredit = 0
                tRs!SaleType = 0
                tRs!Locked = 0
                tRs!CreditlimitCredit = 0
                tRs!CreditlimitCredit = 0
            ElseIf mTableName = "TblItemsUnits" Then
                'tRs!ItemId = mMaxId
                tRs!UnitID = 1
            ElseIf mTableName = "TblEmployee" Then
                tRs.Fields.Item(Index) = mMaxId
            Else
                tRs.Fields.Item(Index) = mMaxId
            End If
            
            
            
        End If
        Dim k As Integer
        For k = 0 To UBound(FieldValue) Step 2
            tRs.Fields.Item(FieldValue(k)).Value = FieldValue(k + 1)
            'Debug.Print FieldValue(k) & " " & tRs.Fields.Item(FieldValue(k)).Value
        Next
        '*************************
        'Debug.Print "fields count " & tRs.Fields.count
        Dim j As Integer
        For j = 0 To tRs.Fields.Count - 1

            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                   If tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = "GroupID" Then
                        j = j
                   End If
                If tRs.Fields.Item(j).Type = adInteger Or tRs.Fields.Item(j).Type = adCurrency Or tRs.Fields.Item(j).Type = adBoolean Or tRs.Fields.Item(j).Type = adSmallInt Or tRs.Fields.Item(j).Type = adBigInt Or tRs.Fields.Item(j).Type = adTinyInt Or tRs.Fields.Item(j).Type = adUnsignedTinyInt Or tRs.Fields.Item(j).Type = adNumeric Or tRs.Fields.Item(j).Type = adDouble Or tRs.Fields.Item(j).Type = adDecimal Then
                    If tRs.Fields.Item(j).Type = adBoolean Then
                        tRs.Fields.Item(j).Value = (UCase(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "TRUE") Or (UCase(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "-1") Or (Val(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = -1)
                    Else
'                        If tGrd.ColComboList(tGrd.ColIndex(tRS.Fields.Item(j).Name)) <> "" Then
'                            tRS.Fields.Item(j).Value = tGrd.ValueMatrix(i, tGrd.ColIndex(tRS.Fields.Item(j).Name))
'                        Else
                            'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                            If Val(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = 0 Then
                                tRs.Fields.Item(j).Value = Null
                            Else
                                tRs.Fields.Item(j).Value = Val(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)))
                            End If
                            'End If
'                        End If
                    End If
                Else
                    If tRs.Fields.Item(j).Type = adDBTimeStamp Or tRs.Fields.Item(j).Type = adDBTime Or tRs.Fields.Item(j).Type = adDBDate Then
                        If Not IsDate(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) Then
                            tRs.Fields.Item(j).Value = Null
                        Else
                            tRs.Fields.Item(j).Value = tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))
                        End If
                    Else
                        'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                        tRs.Fields.Item(j).Value = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))
                        'End If
                    End If
                End If
            End If
'            If tGrd.ColIndex(tRs.Fields.Item(j).Name) = "InsuranceState" Or Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "„ƒ„‰ ⁄·ÌÂ" Or tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = "€Ì— „ƒ„‰ ⁄·ÌÂ" Then
'                If Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & "")) = "„ƒ„‰ ⁄·ÌÂ" Then
'                    tRs.Fields.Item(j).Value = 1
'                Else
'                    tRs.Fields.Item(j).Value = 0
'                End If
'            End If
            'Debug.Print tRs.Fields.Item(j).Name & " = " & tRs.Fields.Item(j).Value
        Next
          If mTableName = "TblItemsUnits" Then
                'tRs!ItemId = mMaxId
                    
                'tRs!UnitID = Val(tGrd.TextMatrix(i, tGrd.ColIndex("UnitId") & ""))
        End If
        
        Dim mNewCode As String
        If UCase(mTableName) = "TBLITEMS" Then
            mNewCode = GetNewCode(Val(tGrd.TextMatrix(i, tGrd.ColIndex("GroupID"))), mTableName)
            tRs!Fullcode = mNewCode
            tRs!itemcode = mNewCode
            tRs!barCodeNO = mNewCode
            tRs!code = mNewCode
            tGrd.TextMatrix(i, tGrd.ColIndex("Fullcode")) = mNewCode
            tGrd.TextMatrix(i, tGrd.ColIndex("code")) = mNewCode
        End If
        If UCase(mTableName) = "TBLCUSTEMERS" Then
            'mNewCode = GetNewCode(Val(tGrd.TextMatrix(i, tGrd.ColIndex("ClassCustomersId"))), mTableName)
            'tRs!Fullcode = mNewCode
            'tRs!code = mNewCode
            
            
            'tGrd.TextMatrix(i, tGrd.ColIndex("Fullcode")) = mNewCode
          '  tGrd.TextMatrix(i, tGrd.ColIndex("code")) = mNewCode
        End If
        
tRs.Update
NextStep:
    Next
    tRs.Close
        DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
    Exit Sub
Err:
    If Err.Number = -2147217887 Then        ' one item is empty
        Resume Next
    End If
    '    Resume Next
End Sub



Private Function GetNewGroupCode(LngParentGroupID As Long, Optional ByVal mTableName As String = "") As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String
    If mTableName = "" Then
        mTableName = "Groups"
    End If
    On Error GoTo ErrTrap
    StrSQL = "Select GroupCode From " & mTableName & "  Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").Value), "", rs("GroupCode").Value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableName & "  Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").Value), "", rs("GroupCode").Value)
        IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function



Private Function GetNewCode(LngParentGroupID As Long, Optional ByVal mTableName As String = "", Optional ByVal mTableGroupName As String = "Groups", Optional ByVal mFieldGroup As String = "GroupID") As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String
    If mTableName = "" Then
        mTableName = "Groups"
    End If
    On Error GoTo ErrTrap
    StrSQL = "Select Max(Code) Code From " & mTableName & "  Where " & mFieldGroup & " =" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("Code").Value), "", rs("Code").Value)
    Else
        StrParentCode = "000"
    End If

     Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim mTmpGroup2  As String
    If Not rs.BOF Then
        StrNewGroupCode = rs!code & ""
        mTmpGroup2 = Replace(StrParentCode, StrNewGroupCode, "")
    End If
    If Trim(mTmpGroup2) = "" Then mTmpGroup2 = "000"
    rs.Close
    Dim mTmp As Long
    mTmp = Val(mTmpGroup2) + 1
    If Len(CStr(mTmp)) = 1 Then
        StrParentCode = "00" & mTmp
    ElseIf Len(CStr(mTmp)) = 2 Then
        StrParentCode = "0" & mTmp
    ElseIf Len(CStr(mTmp)) = 3 Then
        StrParentCode = "" & mTmp
    End If
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From " & mTableGroupName & "   Where GroupID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("Code").Value), "", rs("Code").Value)
        IntTemp = Val(Mid(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrLastGroupCode & StrParentCode
    End If

    rs.Close
    Set rs = Nothing
    GetNewCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function



