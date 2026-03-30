Attribute VB_Name = "ModDatesLib"
Option Explicit

Public Function GetMonthDaysCount(IntMonth As Integer, _
                                  IntYear As Integer) As Integer
    Dim MaxDay As Integer

    Select Case IntMonth

        Case 1, 3, 5, 7, 8, 10, 12
            MaxDay = 31

        Case 2

            If Month(DateAdd("d", 1, CDate("28/2/" & IntYear))) = 2 Then
                MaxDay = 29
            Else
                MaxDay = 28
            End If

        Case Else
            MaxDay = 30
    End Select

    GetMonthDaysCount = MaxDay
End Function

Public Function DisplayDate(ConverDate As Date) As String

    If SystemOptions.UserInterface = ArabicInterface Then
        DisplayDate = Format(ConverDate, "yyyy/Mm/dd")
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        DisplayDate = Format(ConverDate, "dd/Mm/yyyy")
    End If

End Function

Public Function ConvertHoursToMints(StrTime As String) As Long
    Dim VarTemp As Variant
    Dim LngTemp As Long
    VarTemp = Split(StrTime, ":", , vbTextCompare)
    LngTemp = (val(VarTemp(0)) * 60) + val(VarTemp(1))
    ConvertHoursToMints = LngTemp
End Function

Public Function CalculateLateDiscount(IntEmp_ID As Integer, _
                                      StrLateHours As String) As Currency
    Dim RsDis As ADODB.Recordset
    Dim RsEmp As ADODB.Recordset

    Dim StrSQL As String
    Dim i As Integer
    Dim SngOldPart As Single
    Dim SngNewPart  As Single
    Dim SngDayValue As Single
    Dim LngMints As Long
    Dim SngDisValue As Single

    StrSQL = "Select * From tblSliceDiscount Order By Slice_ID"
    Set RsDis = New ADODB.Recordset
    RsDis.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsDis.BOF Or RsDis.EOF Then
        CalculateLateDiscount = 0
        RsDis.Close
        Set RsDis = Nothing
        Exit Function
    Else
        Set RsEmp = New ADODB.Recordset
        StrSQL = "Select tblEmployee.Emp_Salary From tblEmployee Where tblEmployee.Emp_ID=" & IntEmp_ID
        RsEmp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsEmp.BOF Or RsEmp.EOF) Then
            If Not (IsNull(RsEmp("Emp_Salary").value)) Then
                SngDayValue = RsEmp("Emp_Salary").value / 30
            Else
                SngDayValue = 0
            End If
        End If

        RsEmp.Close
        Set RsEmp = Nothing
        LngMints = ConvertHoursToMints(StrLateHours)
    
        RsDis.MoveFirst

        For i = 1 To RsDis.RecordCount
            SngOldPart = RsDis("Late_Time").value
            RsDis.MoveNext

            If Not (RsDis.EOF) Then
                SngNewPart = RsDis("Late_Time").value
            End If

            If LngMints < SngOldPart Then
                SngDisValue = 0
                Exit For
            ElseIf (LngMints >= SngOldPart) And (LngMints <= SngNewPart) Then
                SngDisValue = SngDayValue * RsDis("Dis_Type").value
                Exit For
            ElseIf RsDis.EOF And (LngMints > SngNewPart) Then
                RsDis.MoveLast
                SngDisValue = SngDayValue * RsDis("Dis_Type").value
                Exit For
            End If

            If RsDis.EOF Then
                Exit For
            End If

        Next i

    End If

    SngDisValue = FormatCurrency(SngDisValue, 2)
    CalculateLateDiscount = SngDisValue
End Function

Public Function ConvertMintsToHours(LngMints As Long) As String
    Dim VarTemp As String

    VarTemp = CStr(Format((LngMints \ 60), "00")) & ":" & CStr(Format((LngMints Mod 60), "00"))

    ConvertMintsToHours = VarTemp
End Function

Public Function SQLDate(ConvertDate As Date, _
                        Optional BolPutChar As Boolean = False, _
                        Optional BolPutSep As Boolean = True, Optional BolPutSep2 As Boolean = True) As String

    Dim StrTemp As String
    Dim StrRes As String
    Dim IntMonthB As Integer
    Dim IntMonthE As Integer
    Dim IntDay As Integer
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim StrMonthPrev As String
    Dim StrTempqq As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        IntDay = day(ConvertDate)
        IntMonth = Month(ConvertDate)
        IntYear = year(ConvertDate)
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
        ElseIf BolPutSep2 Then
            StrRes = "" & IntYear & "-" & IntMonth & "-" & (IntDay) & ""
            StrTemp = "'" & StrRes & "'"
        
        End If

    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrRes = Format(ConvertDate, "mm/dd/yyyy")

        If BolPutChar = False Then
            SQLDate = StrRes
        ElseIf BolPutChar = True Then
            StrTemp = "#" & StrRes & "#"
            SQLDate = StrTemp
        End If
    End If

End Function

Public Function GetWeekStartEND(M_Date As Date, _
                                IntType As Integer) As Date
    Dim IntI As Integer
    Dim j  As Integer

    If IntType = 0 Then
        IntI = DatePart("w", M_Date, vbSaturday, vbFirstJan1)
        GetWeekStartEND = M_Date - (IntI - 1)
    ElseIf IntType = 1 Then
        IntI = DatePart("w", M_Date, vbSaturday, vbFirstJan1)
        GetWeekStartEND = M_Date + (7 - IntI)
    End If

End Function
