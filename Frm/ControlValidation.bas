Attribute VB_Name = "ControlValidation"
Private Declare Function StringFromGUID2 _
                          Lib "ole32.dll" (rguid As Any, _
                                           ByVal lpstrClsId As Long, _
                                           ByVal cbMax As Long) As Long
Private Const Click                As String = "Click"
'Collection of validation classes
Private mValidatedTextBoxes        As Collection

Private mValidatedGrids            As Collection
'****************************************
Public mValidatedAutoCompleteGrids As Collection
Const staticName = "txtAuto"

'*****************************************
Private mMenuClass              As Dictionary
'*********************************************
Dim AutoCompleteItemsDictionary As Dictionary
Dim AutoCompleteClassDictionary As Dictionary

Private Function AuControlExists(ByRef ctlName As String, ByRef frm As Form) As Boolean
    On Error GoTo ErrHandler
    Dim ctl As Control
    Set AutoCtrl = Nothing
    For Each ctl In frm.Controls
        If ctl.Name = ctlName Then
            Set AutoCtrl = ctl
            ControlExists = True
            Exit For
        End If
    Next ctl
  
ExitHere:
    Exit Function
ErrHandler:
    Debug.Print Err, Err.description
    Resume ExitHere
End Function

Sub ValidateFormGrid(frm As Object, ByRef RS2 As Variant, ByRef grd As Variant)

    For Each ctl In frm.Controls
        If Not IsArray(ctl) Then
            If TypeOf ctl Is TextBox Then
                If ctl.LinkItem <> "False" Then
                    ctl.LinkItem = "True"
                End If
            ElseIf TypeOf ctl Is vsFlexGrid Then
                If ctl.Tag <> "False" Then
                    ctl.Tag = "True"
                End If
            End If
        End If
    Next

    Dim clsValidateTMP As clsGridValidation    'Temporary class object to be added to collection
    Dim ctlLoopTMP     As Control    'Temporary control object to use when looping through controls
    Set mValidatedGrids = New Collection    'Create a new, empty collection
    For Each ctlLoopTMP In frm.Controls    'Loop through all controls on form
        
        If TypeOf ctlLoopTMP Is vsFlexGrid Then
            If SearchInArray(grd, ctlLoopTMP.Name) <> -1 Then
                If ctlLoopTMP.Tag = "True" Then
                    'Create a new instance of the validation class
                    Set clsValidateTMP = New clsGridValidation
                    'Connect it to the current control
                    Set clsValidateTMP.frmToValidate = frm
                    Dim mIndexRs As Long

                    Set clsValidateTMP.tmpGrdToValidate = grd
                    Set clsValidateTMP.grdToValidate = ctlLoopTMP
                    If ctlLoopTMP.AccessibleDescription <> "" Then
                        For Each ctl In frm.Controls
                            If TypeOf ctl Is vsFlexGrid Then
                                If ctl.Name = ctlLoopTMP.AccessibleDescription Then
                                    Set clsValidateTMP.grdParentToValidate = ctl    '
                                End If
                            End If
                        Next
                    End If
                    mValidatedGrids.Add clsValidateTMP
                    'Add it to the collection
                End If
            End If
        End If
    Next
    'Tidy up object references
    Set clsValidateTMP = Nothing
    Set ctlLoopTMP = Nothing
End Sub
Public Function SearchInArray(ByVal mGrdArray As Variant, ByVal mValue As String) As Long
    On Error Resume Next

    SearchInArray = False

    For i = 0 To UBound(mGrdArray)
        If mValue = Trim(mGrdArray(i)) Then
            SearchInArray = i
            Exit Function
        End If
    Next i
    SearchInArray = -1
End Function

Public Sub SetFormMenu(frm As Form)
       
    Dim clsValidateTMP As cPopupMenu     'Temporary class object to be added to collection
    If mMenuClass Is Nothing Then
        Set mMenuClass = New Dictionary
    End If
    'Set clsValidateTMP.frm = frm
    Set clsValidateTMP = New cPopupMenu
    If IsCtrInForm(frm, "PicToolbar") Then
     
        'Connect it to the current control
        Set clsValidateTMP.ToolbarPictureBox = frm.PicToolbar
     
    End If
    Set clsValidateTMP.CurrentForm = frm
    If Not mMenuClass.Exists(frm.Name) Then
        mMenuClass.Add frm.Name, clsValidateTMP
    End If
    'Tidy up object references
    Set clsValidateTMP = Nothing
    
End Sub



Sub ValidateForm(frm As Form)

    For Each ctl In frm.Controls
        If Not IsArray(ctl) Then
            If TypeOf ctl Is TextBox Then
                If ctl.LinkItem <> "False" Then
                    ctl.LinkItem = "True"
                End If
            ElseIf TypeOf ctl Is vsFlexGrid Then
                If ctl.Tag <> "False" Then
                    ctl.Tag = "True"
                End If
            End If
        End If
    Next

    Dim clsValidateTMP As clsTextValidation    'Temporary class object to be added to collection
    Dim ctlLoopTMP     As Control    'Temporary control object to use when looping through controls
    Set mValidatedTextBoxes = New Collection    'Create a new, empty collection
    'Set clsValidateTMP.frm = frm
    For Each ctlLoopTMP In frm.Controls    'Loop through all controls on form
        If TypeOf ctlLoopTMP Is TextBox Then    'If the current control is a textbox...
            If ctlLoopTMP.LinkItem = "True" And Not IsArray(ctlLoopTMP) Then
                'Create a new instance of the validation class
                Set clsValidateTMP = New clsTextValidation
                'Connect it to the current control
                Set clsValidateTMP.TextBoxToValidate = ctlLoopTMP
                'Add it to the collection
                mValidatedTextBoxes.Add clsValidateTMP
            End If
        ElseIf TypeOf ctlLoopTMP Is CommandButton Then
            If ctlLoopTMP.Tag = Click Then
                'Create a new instance of the validation class
                Set clsValidateTMP = New clsTextValidation
                'Connect it to the current control
                Set clsValidateTMP.CommandButtonToClick = ctlLoopTMP
                'Add it to the collection
                mValidatedTextBoxes.Add clsValidateTMP
            End If
        ElseIf TypeOf ctlLoopTMP Is vsFlexGrid Then
            If ctlLoopTMP.Tag = "True" Then
                'Create a new instance of the validation class
                Set clsValidateTMP = New clsTextValidation
                'Connect it to the current control
                Set clsValidateTMP.grdToValidate = ctlLoopTMP
                'Add it to the collection
                mValidatedTextBoxes.Add clsValidateTMP
            End If
        End If
    Next
    'Tidy up object references
    Set clsValidateTMP = Nothing
    Set ctlLoopTMP = Nothing
End Sub

Sub ValidateFormAutoComplete(frm As Form)
    Exit Sub

    For Each ctl In frm.Controls
        If Not IsArray(ctl) Then
            If TypeOf ctl Is TextBox Then
                If ctl.LinkTopic <> "False|False" Then
                    ctl.LinkTopic = "True|True"
                End If
            End If
        End If
    Next

    Dim clsValidateTMP As clsTextValidation    'Temporary class object to be added to collection
    Dim ctlLoopTMP     As Control    'Temporary control object to use when looping through controls
    Set mValidatedTextBoxes = New Collection    'Create a new, empty collection
    'Set clsValidateTMP.frm = frm
    For Each ctlLoopTMP In frm.Controls    'Loop through all controls on form
        If TypeOf ctlLoopTMP Is TextBox Then    'If the current control is a textbox...
            If ctlLoopTMP.LinkTopic = "True|True" And Not IsArray(ctlLoopTMP) Then
                'Create a new instance of the validation class
                Set clsValidateTMP = New clsTextValidation
                'Connect it to the current control
                Set clsValidateTMP.TxtAutoCompleate = ctlLoopTMP
                'Add it to the collection
                mValidatedTextBoxes.Add clsValidateTMP
                Set clsValidateTMP.frm = frm
            End If
          
        End If
    Next
    'Tidy up object references
    Set clsValidateTMP = Nothing
    Set ctlLoopTMP = Nothing
End Sub

Public Function CheckOldValue(ByVal grd As vsFlexGrid, _
                              ByVal Row As Long, _
                              ByVal Col As Long) As Boolean
    If Trim(grd.Cell(flexcpData, Row, Col)) = Trim(grd.TextMatrix(Row, Col)) And Trim(grd.TextMatrix(Row, Col)) <> "" Then CheckOldValue = False: Exit Function
    CheckOldValue = True
End Function

Public Function ValidateDataMember(frm As Form, NewRecord As Boolean)
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.LinkItem = "True" Then
                If NewRecord Then
                    ctl.DataMember = ""
                Else
                    ctl.DataMember = ctl.Text
                End If
            End If
        End If
    Next

End Function

Public Function OpenGridRs(ByVal SqlStatment As String) As ADODB.Recordset
    Dim rsDummy As New ADODB.Recordset
   
    
    rsDummy.Open SqlStatment, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim tmpRs As New ADODB.Recordset
    Set OpenGridRs = New ADODB.Recordset
    For i = 0 To rsDummy.Fields.count - 1
        tmpRs.Fields.Append rsDummy.Fields.Item(i).Name, rsDummy.Fields.Item(i).Type, rsDummy.Fields(i).DefinedSize, rsDummy.Fields(i).Attributes
    Next

    If rsDummy.EOF And rsDummy.BOF Then tmpRs.Open: GoTo ExitFun
    rsDummy.MoveFirst
    tmpRs.Open
    Do While Not rsDummy.EOF
        tmpRs.AddNew
        For i = 0 To rsDummy.Fields.count - 1
            tmpRs(i) = rsDummy(i)
        Next
        rsDummy.MoveNext
    Loop
    tmpRs.MoveFirst
ExitFun:
    Set OpenGridRs = tmpRs

End Function

Public Function OpenTmpGridRs(ByVal rsDummy As ADODB.Recordset) As ADODB.Recordset
    'Dim rsDummy As New ADODB.Recordset

    Dim tmpRs As New ADODB.Recordset
    Set OpenTmpGridRs = New ADODB.Recordset
    For i = 0 To rsDummy.Fields.count - 1
        tmpRs.Fields.Append rsDummy.Fields.Item(i).Name, rsDummy.Fields.Item(i).Type, rsDummy.Fields(i).DefinedSize, rsDummy.Fields(i).Attributes
    Next

    'If rsDummy.EOF And rsDummy.BOF Then tmpRs.Open: GoTo ExitFun
    tmpRs.Open

ExitFun:
    Set OpenTmpGridRs = tmpRs

End Function
'Public Function CopyTmpGridRs(ByVal rsDummy As ADODB.Recordset) As ADODB.Recordset
'    'Dim rsDummy As New ADODB.Recordset
'
'    Dim tmpRs As New ADODB.Recordset
'    Set tmpRs = OpenTmpGridRs(rsDummy)
'    For i = 0 To rsDummy.Fields.Count - 1
'
'    Next
'
'    'If rsDummy.EOF And rsDummy.BOF Then tmpRs.Open: GoTo ExitFun
'    tmpRs.Open
'
'ExitFun:
'    Set OpenTmpGridRs = tmpRs
'
'End Function
Public Sub ClearTmpRs(ByRef frm As Form, _
                      ByRef grd As vsFlexGrid, _
                      tmpGrd As Variant, _
                      Optional ByVal mStartRow As Long = 1, _
                      Optional ByRef grdParent As vsFlexGrid, _
                      Optional ByRef mParentRow As Long)
    On Error Resume Next
    Dim j As Long
    'Dim rsGrd As New ADODB.Recordset
    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Set rsGrd = frm.rsArray(mIndexRs)
    If rsGrd.EOF And rsGrd.BOF Then Exit Sub
    If rsGrd.EOF Then rsGrd.MoveFirst

    Set rsDummy = OpenTmpGridRs(rsGrd)
    If Not grdParent Is Nothing Then    ' Edit By Samy 6-12-2011 Õ—ÌÂ Ê⁄œ«·Â
        Do While Not rsGrd.EOF
            If rsGrd!ID <> grdParent.TextMatrix(mParentRow, grdParent.ColIndex("SerID")) Then
                rsDummy.AddNew
                For i = 0 To rsGrd.Fields.count - 1
                    rsDummy(i) = rsGrd(i)
                Next
            End If
            rsGrd.MoveNext
        Loop
    End If
    'rsDummy.MoveFirst
    Set rsGrd = rsDummy

    'rsGrd.Update
    ' rsGrd.MoveFirst
    frm.rsArray(mIndexRs) = rsGrd
End Sub
Public Sub ClearTmpRs2(ByRef frm As Form, _
                       ByRef grd As vsFlexGrid, _
                       tmpGrd As Variant, _
                       Optional ByVal mStartRow As Long = 1, _
                       Optional ByRef grdParent As vsFlexGrid, _
                       Optional ByRef mParentRow As Long)
    On Error Resume Next
    Dim j As Long
    'Dim rsGrd As New ADODB.Recordset
    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Set rsGrd = frm.rsArray(mIndexRs)
    If rsGrd.EOF And rsGrd.BOF Then Exit Sub
    If rsGrd.EOF Then rsGrd.MoveFirst

    Dim rsDummy2 As New ADODB.Recordset

    Set rsDummy = OpenTmpGridRs(rsGrd)
    If Not grdParent Is Nothing Then    ' Edit By Samy 6-12-2011 Õ—ÌÂ Ê⁄œ«·Â
        Do While Not rsGrd.EOF
            If rsGrd!ID = grdParent.TextMatrix(mParentRow, grdParent.ColIndex("SerID")) Then
                rsGrd.Delete
                '                For i = 0 To rsGrd.Fields.Count - 1
                '                    rsDummy(i) = rsGrd(i)
                '                Next
            End If
            rsGrd.MoveNext
        Loop
    End If
    Set rsDummy2 = rsGrd

    rsGrd.MoveFirst
    'rsDummy.MoveFirst
    ' Set rsGrd = rsDummy

    'rsGrd.Update
    ' rsGrd.MoveFirst
    frm.rsArray(mIndexRs) = rsGrd
End Sub

Public Sub SaveTmpRs(ByRef frm As Form, _
                     ByRef grd As vsFlexGrid, _
                     tmpGrd As Variant, _
                     Optional ByVal mStartRow As Long = 1, _
                     Optional ByRef grdParent As vsFlexGrid, _
                     Optional ByVal mParentRow As Long, _
                     Optional ByVal CheckPoint As String = "", _
                     Optional ByVal arrIndex As Integer = -1)
    Dim j As Long
    If arrIndex = -1 Then
        mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Else
        mIndexRs = arrIndex
    End If
    
    Set rsGrd = frm.rsArray(mIndexRs)
    
    For i = mStartRow To grd.Rows - 1
        mRow = i
        If mRow <> 0 Then
            mCheckPoint = ""
            If (grd.ColIndex(CheckPoint) <> -1 And CheckPoint <> "") Then
                'If i > Grd.Rows - 1 Then
                mCheckPoint = Trim(grd.TextMatrix(i, grd.ColIndex(CheckPoint)))
            End If

            If CheckPoint = "" Or mCheckPoint <> "" Then
                If grd.ColIndex("SerID") <> -1 Then
                    If Trim(grd.TextMatrix(mRow, grd.ColIndex("SerID"))) = "" Then
                        grd.TextMatrix(mRow, grd.ColIndex("SerID")) = GetNewID
                    End If
                End If
                If grd.ColIndex("ID") <> -1 Then
                    If Trim(grd.TextMatrix(mRow, grd.ColIndex("ID"))) = "" And grd.AccessibleDescription = "" Then
                        grd.TextMatrix(mRow, grd.ColIndex("ID")) = GetNewID
                    ElseIf grd.AccessibleDescription <> "" Then
                        If mParentRow = 0 Then grd.Rows = 1: grd.Rows = 2
                        grd.TextMatrix(mRow, grd.ColIndex("ID")) = grdParent.TextMatrix(mParentRow, grdParent.ColIndex("SerID"))
                    End If
                    If Not rsGrd.EOF Or Not rsGrd.BOF Then
                        rsGrd.MoveFirst
                    End If
                    rsGrd.Find "SerID = '" & Trim(grd.TextMatrix(mRow, grd.ColIndex("SerID"))) & "'"
                    If rsGrd.EOF Then
                        rsGrd.AddNew
                    End If
                End If

                '-----------------------------------------------------
                For j = 0 To rsGrd.Fields.count - 1
                    If grd.ColIndex(rsGrd.Fields.Item(j).Name) <> -1 Then
                        If rsGrd.Fields.Item(j).Name = "Passed" Or rsGrd.Fields.Item(j).Name = "DownPayment" Then
                            xcxcx = 33
                        End If
                        If rsGrd.Fields.Item(j).Type = adInteger Or rsGrd.Fields.Item(j).Type = adCurrency Or rsGrd.Fields.Item(j).Type = adBoolean Or rsGrd.Fields.Item(j).Type = adSmallInt Or rsGrd.Fields.Item(j).Type = adBigInt Or rsGrd.Fields.Item(j).Type = adTinyInt Or rsGrd.Fields.Item(j).Type = adUnsignedTinyInt Or rsGrd.Fields.Item(j).Type = adNumeric Or rsGrd.Fields.Item(j).Type = adUnsignedTinyInt Or rsGrd.Fields.Item(j).Type = adDouble Then
                            If rsGrd.Fields.Item(j).Type = adBoolean Then
                                If (UCase(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "TRUE") Or myRound((grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "1") Or (grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name)) = "-1") Then
                                    rsGrd.Fields.Item(j).Value = 1
                                ElseIf myRound(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "0" Or (UCase(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "FALSE") Then
                                    rsGrd.Fields.Item(j).Value = 0
                                End If
                            Else
                                rsGrd.Fields.Item(j).Value = myRound(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name)))
                            End If
                        Else
                            If rsGrd.Fields.Item(j).Type = adDBTimeStamp Or rsGrd.Fields.Item(j).Type = adDBTime Or rsGrd.Fields.Item(j).Type = adDBDate Then
                                If Not IsDate(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) Then
                                    ' rsGrd.Fields.Item(j).Value = "" ' Date
                                Else
                                    rsGrd.Fields.Item(j).Value = grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))
                                End If
                            Else
                                If Len(Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & ""))) > 100 Then
                                    '  rsGrd.Fields.Item(j).Value = Left(Trim(Grd.TextMatrix(mRow, Grd.ColIndex(rsGrd.Fields.Item(j).Name) & "")), 100)
                                    rsGrd.Fields.Item(j).Value = Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & ""))
                                Else
                                    If rsGrd.Fields.Item(j).Type = adSingle Then
                                        rsGrd.Fields.Item(j).Value = myRound(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & ""))
                                    Else
                                        rsGrd.Fields.Item(j).Value = Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & ""))
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    If arrIndex = -1 Then
        mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Else
        mIndexRs = arrIndex
    End If
    frm.rsArray(mIndexRs) = rsGrd
End Sub

Public Sub SaveRs(ByVal rsDummy As ADODB.Recordset, _
                  ByVal ChekPoint As String, _
                  ByVal Index As String, _
                  ByVal mfldPrimName As Variant, _
                  ByVal mfldPrimValue As Variant, _
                  mTableName As String)
    If rsDummy.EOF And rsDummy.BOF Then Exit Sub
    Cond = "1 =1"
    For i = 0 To UBound(mfldPrimName)
        If UCase(mfldPrimName(i)) = UCase("IDRef") Or UCase(mfldPrimName(i)) = UCase("Id") Or UCase(mfldPrimName(i)) = UCase("NoteType") Then
            Cond = Cond & " And " & mfldPrimName(i) & " =" & mfldPrimValue(i)
        Else
            Cond = Cond & " And " & mfldPrimName(i) & " = N'" & mfldPrimValue(i) & "'"
        End If
    Next
    ss = "Delete from " & mTableName & " Where " & Cond
    '    ss = DoAcc(ss)
    db.Execute ss
    
    s = "Select * from " & mTableName
    Dim RS As New ADODB.Recordset
    Set RS = OpenRecordSet(s, adOpenKeyset, adLockOptimistic)
    rsDummy.MoveFirst
    Do While Not rsDummy.EOF
        If ChekPoint <> "" Then
            If Trim(rsDummy(ChekPoint) & "") = "" Then GoTo nextStep
            If rsDummy.Fields.Item(ChekPoint).Type = adInteger Then
                If myRound(rsDummy(ChekPoint) & "") = 0 Then GoTo nextStep
            End If
        End If
        mmID2 = rsDummy!ID
        mmSerID2 = rsDummy!SerID
        '**********************
        RS.AddNew
        For i = 0 To rsDummy.Fields.count - 1

            For j = 0 To RS.Fields.count - 1

                If UCase(Trim(RS.Fields(j).Name)) = UCase(Trim(rsDummy.Fields(i).Name)) Then
                    If RS.Fields.Item(j).Name = "OrderDate" Then
                        xcxcx = 33
                    End If
                    If RS.Fields.Item(j).Type = adInteger Or RS.Fields.Item(j).Type = adCurrency Or RS.Fields.Item(j).Type = adBoolean Or RS.Fields.Item(j).Type = adSmallInt Or RS.Fields.Item(j).Type = adBigInt Or RS.Fields.Item(j).Type = adTinyInt Or RS.Fields.Item(j).Type = adUnsignedTinyInt Or RS.Fields.Item(j).Type = adNumeric Or RS.Fields.Item(j).Type = adUnsignedTinyInt Or RS.Fields.Item(j).Type = adDouble Then
                        If RS.Fields.Item(j).Type = adBoolean Then
                            RS.Fields.Item(j).Value = IIf(rsDummy.Fields.Item(i).Value & "" = "", False, rsDummy.Fields.Item(i).Value)
                        Else
                            RS.Fields.Item(j).Value = myRound(rsDummy.Fields.Item(i).Value & "")
                        End If
                    Else
                        If RS.Fields.Item(j).Type = adDBTimeStamp Or RS.Fields.Item(j).Type = adDBTime Or RS.Fields.Item(j).Type = adDBDate Then
                            If Not IsDate(rsDummy.Fields.Item(i).Value) Then
                                RS.Fields.Item(j).Value = Null
                            Else
                                RS.Fields.Item(j).Value = rsDummy.Fields.Item(i).Value
                            End If
                        Else
                            If RS.Fields.Item(j).Type = adGUID Then
                                If rsDummy.Fields.Item(i).Value & "" = "" Then
                                    RS.Fields.Item(j).Value = CreateGUID("")
                                Else
                                    RS.Fields.Item(j).Value = Trim(rsDummy.Fields.Item(i).Value)
                                End If
                            Else    ' string
                                If Len(Trim(rsDummy.Fields.Item(i).Value)) > 100 Then
                                    ' rs.Fields.Item(j).Value = Left(Trim(rsDummy.Fields.Item(i).Value), 100)
                                    RS.Fields.Item(j).Value = Trim(rsDummy.Fields.Item(i).Value)
                                Else
                                    RS.Fields.Item(j).Value = Trim(rsDummy.Fields.Item(i).Value)
                                End If
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next
        Next
      
        If mmId <> mmID2 Then
            xx = 4
            II = 0
        End If
        II = II + 1
        If Index <> "" Then RS(Index) = II
        For m = 0 To UBound(mfldPrimName)
            RS(mfldPrimName(m)) = mfldPrimValue(m)
        Next
UpdateRecordSet RS
nextStep:
        rsDummy.MoveNext
        If Not rsDummy.EOF Then
            mmId = rsDummy!ID
            mmSerID = rsDummy!SerID
        End If
    Loop
End Sub




Public Function CreateGUID(Optional strRemoveChars As String = "{}-") As String

    Dim udtGUID   As Guid
    Dim strGUID   As String
    Dim bytGUID() As Byte
    Dim lngLen    As Long
    Dim lngRetVal As Long
    Dim lngPos    As Long

    lngLen = 40
    bytGUID = String(lngLen, 0)

    CoCreateGuid udtGUID

    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
    strGUID = bytGUID
    If (Asc(mId$(strGUID, lngRetVal, 1)) = 0) Then
        lngRetVal = lngRetVal - 1
    End If

    strGUID = left$(strGUID, lngRetVal)

    For lngPos = 1 To Len(strRemoveChars)
        strGUID = Replace(strGUID, mId(strRemoveChars, lngPos, 1), "")
    Next
    CreateGUID = strGUID
End Function


