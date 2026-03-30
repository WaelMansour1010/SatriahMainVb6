VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

'The important bit - note the WithEvents keyword.
Dim mIndexRs As Long
Private WithEvents grd As vsFlexGrid
Attribute grd.VB_VarHelpID = -1
Private WithEvents grdParent As vsFlexGrid
Attribute grdParent.VB_VarHelpID = -1



'*********************************
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1
Public rsGrd As ADODB.Recordset
Private mParentLineNo As String

Private tmpRs() As New ADODB.Recordset
Private tmpGrd As Variant

Public Property Set tmpRsToValidate(ByRef SetRs As Variant)
    ReDim tmpRs(UBound(SetRs))
    For i = 0 To UBound(SetRs)
        Set tmpRs(i) = SetRs(i)
    Next
End Property

Public Property Set tmpGrdToValidate(ByRef SetGrd As Variant)
    ReDim tmpGrd(UBound(SetGrd))
    For i = 0 To UBound(SetGrd)
        tmpGrd(i) = SetGrd(i)
    Next
End Property

Public Property Set frmToValidate(SetForm As Form)
'Setting this property connects the control on the form to this instance of this class
    Set frm = SetForm
End Property

Public Property Set grdToValidate(SetGrd As vsFlexGrid)
'Setting this property connects the control on the form to this instance of this class
    Set grd = SetGrd
End Property

Public Property Set grdParentToValidate(SetGrd As vsFlexGrid)
'Setting this property connects the control on the form to this instance of this class
    Set grdParent = SetGrd
End Property

'Public Property Set RsToValidate(SetRs As ADODB.Recordset)
'    'Setting this property connects the control on the form to this instance of this class
'    Set rsGrd = SetRs
'End Property

Private Sub FillRs(ByVal mRow As Long)
    Dim j As Long

    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Set rsGrd = frm.rsArray(mIndexRs)
    If rsGrd.EOF Then
        rsGrd.AddNew
    End If
    '-----------------------------------------------------
    For j = 0 To rsGrd.Fields.count - 1
        If grd.ColIndex(rsGrd.Fields.Item(j).Name) <> -1 Then
            If rsGrd.Fields.Item(j).Type = adInteger Or rsGrd.Fields.Item(j).Type = adCurrency Or rsGrd.Fields.Item(j).Type = adBoolean Or rsGrd.Fields.Item(j).Type = adSmallInt Or rsGrd.Fields.Item(j).Type = adBigInt Or rsGrd.Fields.Item(j).Type = adTinyInt Or rsGrd.Fields.Item(j).Type = adUnsignedTinyInt Or rsGrd.Fields.Item(j).Type = adNumeric Or rsGrd.Fields.Item(j).Type = adUnsignedTinyInt Or rsGrd.Fields.Item(j).Type = adDouble Then
                If rsGrd.Fields.Item(j).Type = adBoolean Then
                    rsGrd.Fields.Item(j).Value = (UCase(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "TRUE") Or (UCase(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) = "-1")
                Else
                    rsGrd.Fields.Item(j).Value = val(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name)))
                End If
            Else
                If rsGrd.Fields.Item(j).Type = adDBTimeStamp Or rsGrd.Fields.Item(j).Type = adDBTime Or rsGrd.Fields.Item(j).Type = adDBDate Then
                    If Not IsDate(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))) Then

                        rsGrd.Fields.Item(j).Value = Null
                    Else
                        rsGrd.Fields.Item(j).Value = grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name))
                    End If
                Else
                    If Len(Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & ""))) > 100 Then

                        rsGrd.Fields.Item(j).Value = (Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & "")))

                        ' rsGrd.Fields.Item(j).Value = Left(Trim(Grd.TextMatrix(mRow, Grd.ColIndex(rsGrd.Fields.Item(j).Name) & "")), 100)

                    Else
                    
                        'rsGrd.Fields.Item(j).Value = (Trim(Grd.TextMatrix(mRow, Grd.ColIndex(rsGrd.Fields.Item(j).name) & "")))
                        If rsGrd.Fields.Item(j).Type = adSingle Then
                            rsGrd.Fields.Item(j).Value = (val(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & "")))
                        Else
                            rsGrd.Fields.Item(j).Value = (Trim(grd.TextMatrix(mRow, grd.ColIndex(rsGrd.Fields.Item(j).Name) & "")))
                        End If
                    End If
                End If
            End If
        End If
    Next

    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    frm.rsArray(mIndexRs) = rsGrd

End Sub

Public Sub FillGrid(ByRef mGrd As vsFlexGrid)
    mGrd.Rows = 1
    mGrd.Rows = 2
    Dim i As Integer, j As Integer
    i = 1

    mIndexRs = SearchInArray(tmpGrd, mGrd.Name)
    Dim rsDummy As New ADODB.Recordset
    If mIndexRs = 2 Then
        X = 3
    End If

    Set rsDummy = frm.rsArray(mIndexRs)

    If rsDummy.EOF And rsDummy.BOF Then
        'rsDummy.MoveLast
        'rsDummy.MoveNext
        Exit Sub
    End If
    On Error Resume Next
    rsDummy.Find "ID = '" & Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID"))) & "'", 0, adSearchForward, 1
    Do While Not rsDummy.EOF
        If Trim(rsDummy!ID) = Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID"))) Then
            mGrd.Rows = mGrd.Rows + 1
            For j = 0 To rsDummy.Fields.count - 1
                If mGrd.ColIndex(rsDummy.Fields.Item(j).Name) <> -1 Then
                    If rsDummy.Fields.Item(j).Type = adCurrency Then
                        mGrd.TextMatrix(i, mGrd.ColIndex(rsDummy.Fields.Item(j).Name)) = MyFormat(val(rsDummy.Fields.Item(j) & ""))
                    Else
                        mGrd.TextMatrix(i, mGrd.ColIndex(rsDummy.Fields.Item(j).Name)) = Trim(rsDummy.Fields.Item(j) & "")
                    End If
                End If
            Next
            i = i + 1
        End If
        rsDummy.MoveNext
    Loop
    GridSerial mGrd
    On Error Resume Next
    CallByName frm, mGrd.Name & "_AfterGridFill", VbMethod
End Sub

Private Sub DeleteChildRs(ByVal mGrd As vsFlexGrid, ByVal mSerID As String)
    Dim i As Integer, j As Integer
    i = 1

    mIndexRs = SearchInArray(tmpGrd, mGrd.Name)
    '------------------------------------------
    mGrd.Rows = 1
    mGrd.Rows = 2
    Dim rsDummy As New ADODB.Recordset
    Set rsDummy = frm.rsArray(mIndexRs)
    If rsDummy.EOF And rsDummy.BOF Then Exit Sub

    rsDummy.Find "ID = '" & Trim(mSerID) & "'", 0, adSearchForward, 1
    Do While Not rsDummy.EOF
        If Trim(rsDummy!ID) = Trim(mSerID) Then
            mChildID = Trim(rsDummy!SerID)
            rsDummy.Delete adAffectCurrent
            Dim ctl As Control
            For Each ctl In frm.Controls
                If TypeOf ctl Is vsFlexGrid Then
                    If ctl.AccessibleDescription = mGrd.Name Then
                        DeleteChildRs ctl, mChildID
                        mIndexRs = SearchInArray(tmpGrd, mGrd.Name)
                    End If
                End If
            Next
        End If
        rsDummy.MoveNext
    Loop
    frm.rsArray(mIndexRs) = rsDummy
End Sub

Public Sub Multi_Find(ByRef tmpRs As ADODB.Recordset, sCriteria As String)
    Dim rsClone As ADODB.Recordset
    Set rsClone = tmpRs.Clone
    rsClone.filter = sCriteria
    If rsClone.EOF Or rsClone.BOF Then
        tmpRs.MoveLast
        tmpRs.MoveNext
    Else
        tmpRs.Bookmark = tmpRs.Bookmark
    End If
    rsClone.Close
    Set rsClone = Nothing
End Sub

Private Function SearchInArray(ByVal mArray As Variant, ByVal mValue As String) As Long
    SearchInArray = False
    For i = 0 To UBound(mArray)
        If mValue = Trim(mArray(i)) Then
            SearchInArray = i
            Exit Function
        End If
    Next i
End Function

Private Sub ExcelSheet_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, _
                                          ByVal SaveAsUI As Boolean, _
                                          Cancel As Boolean)
    MsgBox "Samy"
End Sub

Private Sub excelWB_BeforeClose(Cancel As Boolean)
    MsgBox ""
End Sub

Private Sub Grd_AfterRowColChange(ByVal OldRow As Long, _
                                  ByVal OldCol As Long, _
                                  ByVal NewRow As Long, _
                                  ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If grd.Rows > 1 Then   ' ›ﬁÿ ⁄‰œ„« Ì €Ì— «·”ÿ—
            mParentLineNo = Trim(grd.TextMatrix(NewRow, grd.ColIndex("SerID")))
            Dim ctl As Control
            For Each ctl In frm.Controls
                If TypeOf ctl Is vsFlexGrid Then
                    If ctl.AccessibleDescription = grd.Name Then
                        FillGrid ctl
                    End If
                End If
            Next
        End If
    End If
End Sub

Private Sub Grd_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Set rsGrd = frm.rsArray(mIndexRs)
    If Trim(grd.TextMatrix(Row, grd.ColIndex("SerID"))) = "" Then
        grd.TextMatrix(Row, grd.ColIndex("SerID")) = GetNewID
    End If
    rsGrd.Find "SerID = '" & Trim(grd.TextMatrix(Row, grd.ColIndex("SerID"))) & "'", 0, adSearchForward, 1
    FillRs Row
End Sub

Private Sub Grd_Click()
    If grd.ColIndex("SerID") <> -1 Then
        If Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID"))) = "" Then
            grd.TextMatrix(grd.Row, grd.ColIndex("SerID")) = GetNewID
        End If
    End If
    If grd.ColIndex("ID") <> -1 Then
        If Trim(grd.TextMatrix(grd.Row, grd.ColIndex("ID"))) = "" And grd.AccessibleDescription = "" Then
            grd.TextMatrix(grd.Row, grd.ColIndex("ID")) = GetNewID
        ElseIf grd.AccessibleDescription <> "" And InStr(grd.AccessibleDescription, ",") = False Then

            If grdParent.Row = 0 Then grd.Rows = 1: grd.Rows = 2
            grd.TextMatrix(grd.Row, grd.ColIndex("ID")) = grdParent.TextMatrix(grdParent.Row, grdParent.ColIndex("SerID"))
        End If
        mParentLineNo = Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID")))
        Dim ctl As Control
        For Each ctl In frm.Controls
            If TypeOf ctl Is vsFlexGrid Then
                If ctl.AccessibleDescription = grd.Name Then
                    FillGrid ctl
                End If

            End If
        Next
    End If

End Sub

Private Sub Grd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo EH

    '*****************************************************************************
    If InStr(grd.AccessibleDescription, ",") Then

        '        SplitPoint = InStr(Grd.AccessibleDescription, ",")
        '        mInsert = IIf(left(Grd.AccessibleDescription, SplitPoint - 1) = "True", True, False)
        '        mDelete = IIf(right(Grd.AccessibleDescription, Len(Grd.AccessibleDescription) - SplitPoint) = "True", True, False)
        '        GridKeyDown Grd, KeyCode, Shift, mInsert, mDelete

        Exit Sub

    End If

    If grd.AccessibleName = "EntryBenfits3" Or grd.AccessibleName = "EntryDues3" Or grd.AccessibleName = "EntryDuesOut3" Or grd.AccessibleName = "EntryDuesInDelete3" Or grd.AccessibleName = "EntryDuesInDelete4" Then
        If grd.Parent.Name = "frmEntryBenefitsPost" Then
        ElseIf grd.Parent.Name = "frmEntryBenefitsDelete2" Then
            GridKeyDown grd, KeyCode, Shift, True, False
            Exit Sub
        Else
            GridKeyDown grd, KeyCode, Shift, True, True
            Exit Sub
        End If
        Exit Sub
    ElseIf grd.AccessibleName = "EntryBenfits4" Or grd.AccessibleName = "EntryDues4" Or grd.AccessibleName = "EntryDuesOut4" Or grd.AccessibleName = "fPCWorksTypesAnalysisExpenses3" Then
        If KeyCode = vbKeyReturn Then
            GridKeyDown grd, KeyCode, Shift, False, True
            Exit Sub
        Else
            If grd.Row = 1 Then
                GridKeyDown grd, KeyCode, Shift, True, True
                Exit Sub
            ElseIf grd.Parent.Name = "frmEntryBenefitsPost" Then
                GridKeyDown grd, KeyCode, Shift, True, True
                Exit Sub
            End If
        End If
    End If

    '        '********************************************************************
    'normalCase:
    OldKey = KeyCode: oldShift = Shift
    If KeyCode = vbKeyDelete And Shift = 1 Then    ' Shift +Delete = Delete the row

    ElseIf KeyCode = vbKeyInsert Then    ' «÷«›… ”ÿ— ÃœÌœ
        If grd.ColIndex("SerID") <> -1 Then
            If Trim(grd.TextMatrix(grd.Row, grd.ColIndex("ID"))) = "" Then
                grd.TextMatrix(grd.Row, grd.ColIndex("SerID")) = GetNewID
            End If
        End If

        If grd.ColIndex("ID") <> -1 Then
            If Trim(grd.TextMatrix(grd.Row, grd.ColIndex("ID"))) = "" And grd.AccessibleDescription = "" Then
                grd.TextMatrix(grd.Row, grd.ColIndex("ID")) = GetNewID
            ElseIf grd.AccessibleDescription <> "" Then
                If grdParent.Row = 0 Then grd.Rows = 1: grd.Rows = 2
                grd.TextMatrix(grd.Row, grd.ColIndex("ID")) = grdParent.TextMatrix(grdParent.Row, grdParent.ColIndex("SerID"))
            End If
        End If
    End If
    ' ********************
    If (OldKey = vbKeyDelete And oldShift = 1) Then    ' Shift +Delete = Delete the row
        Dim rsDummy As ADODB.Recordset
        mIndexRs = SearchInArray(tmpGrd, grd.Name)
        Set rsGrd = frm.rsArray(mIndexRs)

        mSerID = Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID")))
        rsGrd.Find "SerID = '" & Trim(grd.TextMatrix(grd.Row, grd.ColIndex("SerID"))) & "'", 0, adSearchForward, 1
        If Not rsGrd.EOF Then rsGrd.Delete adAffectCurrent
        
         frm.rsArray(mIndexRs) = rsGrd
        Dim ctl As Control
        For Each ctl In frm.Controls
            If TypeOf ctl Is vsFlexGrid Then
                If ctl.AccessibleDescription = grd.Name Then
                    DeleteChildRs ctl, mSerID
                End If
            End If
        Next
    End If
    GridKeyDown grd, KeyCode, Shift, False, False
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Grd_AfterRowColChange 0, 0, grd.Row, grd.Col
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' *******************************
    Exit Sub
EH:
    MyMsgbox MyErrorHandler(Err)
End Sub

Private Sub Grd_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mIndexRs = SearchInArray(tmpGrd, grd.Name)
    Set rsGrd = frm.rsArray(mIndexRs)
    If grd.ColIndex("SerID") <> -1 Then
        If Trim(grd.TextMatrix(Row, grd.ColIndex("SerID"))) = "" Then
            grd.TextMatrix(Row, grd.ColIndex("SerID")) = GetNewID
        End If
    End If
    If grd.ColIndex("ID") <> -1 Then
        If Trim(grd.TextMatrix(Row, grd.ColIndex("ID"))) = "" And grd.AccessibleDescription = "" Then
            grd.TextMatrix(Row, grd.ColIndex("ID")) = GetNewID
        ElseIf grd.AccessibleDescription <> "" Then
            If grdParent.Row = 0 Then grd.Rows = 1: grd.Rows = 2
            grd.TextMatrix(Row, grd.ColIndex("ID")) = grdParent.TextMatrix(grdParent.Row, grdParent.ColIndex("SerID"))
        End If
        '   Multi_Find rsGrd, "ID = " & grd.TextMatrix(Row, grd.ColIndex("ID")) & " And SerID = " & grd.TextMatrix(Row, grd.ColIndex("SerID"))
        'Multi_Find rsGrd, "SerID = '" & Trim(grd.TextMatrix(Row, grd.ColIndex("SerID"))) & "'"  ' And ID = '" & Trim(grd.TextMatrix(Row, grd.ColIndex("ID"))) & "'"
        rsGrd.Find "SerID = '" & Trim(grd.TextMatrix(Row, grd.ColIndex("SerID"))) & "'", 0, adSearchForward, 1
        FillRs Row
    End If

    On Error Resume Next

End Sub

Private Sub Grd_KeyDownEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            KeyCode As Integer, _
                            ByVal Shift As Integer)
    Grd_KeyDown KeyCode, Shift
    On Error Resume Next

End Sub




