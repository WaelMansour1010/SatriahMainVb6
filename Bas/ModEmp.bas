Attribute VB_Name = "ModEmp"
Option Explicit

Public Enum FormatList
    AllType = 0
    StrOnly = 1
    NumOnly = 2
    DateOnly = 3
    CurOnly = 4
    ChrOnly = 5
End Enum

Public Function DataFormat(ByVal StrFormat As FormatList, _
                           ByVal ascii As Long) As Long
    Dim strString As String
    strString = "0123456789"
    DataFormat = ascii

    If ascii = 8 Or ascii = 13 Then Exit Function
    DataFormat = 0

    Select Case StrFormat

        Case AllType '************************* All type *******************************

        Case StrOnly  '************************* Str type *******************************
            strString = strString & ",.;'|:¡*^%$#@!ø)([]{}÷×-+&_"

            If InStr(1, strString, Chr(ascii), vbTextCompare) Then
                DataFormat = 0
            Else
                DataFormat = ascii
            End If

        Case NumOnly '************************* Num type *******************************

            If InStr(1, strString, Chr(ascii), vbTextCompare) Then DataFormat = ascii

        Case DateOnly  '************************* Date type *******************************
            strString = strString & "/-"

            If InStr(1, strString, Chr(ascii), vbTextCompare) Then DataFormat = ascii

        Case CurOnly  '************************* Cur type *******************************
            strString = strString & "."

            If InStr(1, strString, Chr(ascii), vbTextCompare) Then DataFormat = ascii

        Case ChrOnly  '************************* Chr type *******************************
            strString = ",.;'|:¡*^%$#@!ø)([]{}÷×-+&_"

            If InStr(1, strString, Chr(ascii), vbTextCompare) Then
                DataFormat = 0
            Else
                DataFormat = ascii
            End If

    End Select

End Function

Public Function RoundValue(ByVal Number As Single, _
                           RoundCount As Integer) As Single
    Dim RndNum As String
    Dim RoundNum As String
    Dim I As Integer
    RoundValue = 0
    RoundNum = ""

    For I = 1 To RoundCount
        RoundNum = RoundNum & 0
    Next

    RoundNum = "." & RoundNum
    RoundValue = Format(Number, RoundNum)
End Function

Public Function IsRecExist(ByVal StrTable As String, _
                           ByVal StrField As String, _
                           ByVal StrVal As String, _
                           Optional FieldRtn As String, _
                           Optional SQLWhere As String) As String
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RsExist As New ADODB.Recordset
    Dim StrWhere As String

    If SQLWhere <> "" Then SQLWhere = " and " & SQLWhere
    StrWhere = " where " & StrField & "='" & CStr(StrVal) & "'" & SQLWhere
    'StrWhere = " where " & StrField & "='" & CStr(StrVal) & "'" & SqlWhere
    IsRecExist = ""
    My_SQL = "select * From " & StrTable & StrWhere
    RsExist.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly

    If RsExist.RecordCount > 0 Then
        IsRecExist = "True"

        If FieldRtn <> "" Then
            IsRecExist = IIf(IsNull(RsExist.Fields(FieldRtn)), "", RsExist.Fields(FieldRtn))
        End If
    End If

    RsExist.Close
    Set RsExist = Nothing
ErrTrap:
End Function

Public Function FillListBox(ByRef List As ListBox, _
                            ByVal StrSQL As String)
    On Error GoTo ErrTrap
    Dim RsRec As New ADODB.Recordset
    Dim I As Long
    List.Clear
    RsRec.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    If RsRec.RecordCount > 0 Then
        RsRec.MoveFirst

        Do Until RsRec.EOF
            List.AddItem IIf(IsNull(RsRec.Fields(1)), "", RsRec.Fields(1))
            List.ItemData(List.NewIndex) = IIf(IsNull(RsRec.Fields(0)), "", RsRec.Fields(0))
            RsRec.MoveNext
        Loop

    End If

    RsRec.Close
    Set RsRec = Nothing
ErrTrap:
End Function

Public Sub FillFlexGrid(My_Flex As VSFlex8Ctl.vsFlexGrid, _
                        ByVal Col As Integer, _
                        My_SQL As String)
    Dim rs As New ADODB.Recordset
    Dim StrOrder As String
    Dim intStart As Integer
    Dim IntEnd As Integer
    On Error GoTo ErrorHandler
    'My_Flex.ListField = ""
    'My_Flex.BoundText = ""
    'My_Flex.BoundColumn = ""
    intStart = InStr(1, My_SQL, ",", vbTextCompare)
    IntEnd = InStr(1, My_SQL, "From", vbTextCompare)
    StrOrder = Trim(Mid(My_SQL, intStart + 1, IntEnd - intStart - 1)) 'Left(My_SQL, IntEnd - 1)
    Dim TxtCombo As String
    Set rs = Nothing
    StrOrder = " order by " & StrOrder

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then Exit Sub

    My_Flex.Tag = My_SQL
    rs.Open My_SQL & StrOrder, Cn, adOpenStatic, adLockReadOnly

    If rs.RecordCount >= 0 Then
        rs.MoveFirst

        With My_Flex

            Do Until rs.EOF
                TxtCombo = TxtCombo + "#" & IIf(IsNull(rs.Fields(0).value), "", rs.Fields(0).value) & ";" & IIf(IsNull(rs.Fields(1).value), "", rs.Fields(1).value) & "|"
            
                rs.MoveNext
            Loop

            .ColComboList(Col) = TxtCombo
        End With

    Else
    End If

Exit_Sub:
    'Rs.Close
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

