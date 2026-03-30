Attribute VB_Name = "ModReportsFunction_n"

Public Function DefindPiasterType(StrThePiasterValue As String) As String
    ''Â–Â «·œ«·… „‘«»Â…  „«„« „À· DefindPoundeType
    'Dim StrRS As ADODB.Recordset
    'Set StrRS = New ADODB.Recordset
    'StrRS.Open "Options", cn, adOpenStatic, adLockReadOnly, adCmdTable
    'If StrThePiasterValue = "0" Then
    '    DefindPiasterType = ""
    '
    'ElseIf StrThePiasterValue = "1" Then
    '    If IsNull(StrRS("small_currency_unite").Value) = False Then
    '        DefindPiasterType = StrRS("small_currency_unite").Value
    '    Else
    '        DefindPiasterType = ""
    '    End If
    '    'DefindPiasterType = " ﬁ—‘ "
    'ElseIf StrThePiasterValue = " 2" Then
    '    If IsNull(StrRS("small_currency_unite_do").Value) = False Then
    '        DefindPiasterType = StrRS("small_currency_unite_do").Value
    '    Else
    '        DefindPiasterType = ""
    '    End If
    'ElseIf (3 <= Val(StrThePiasterValue) And Val(StrThePiasterValue) <= 10) Then
    '    If IsNull(StrRS("small_currency_unite_sum")) = False Then
    '        DefindPiasterType = StrRS("small_currency_unite_sum").Value
    '    Else
    '        DefindPiasterType = ""
    '    End If
    '
    'Else
    '    If IsNull(StrRS("small_currency_unite_exp")) = False Then
    '        DefindPiasterType = StrRS("small_currency_unite_exp").Value
    '    Else '
    '        DefindPiasterType = ""
    '    End If
    'End If
    'StrRS.Close
    'Set StrRS = Nothing
End Function

Public Function DefindPoundeType(DblThePoundValue As Double) As String
    ''›«∆œ… Â–Â «·œ«·… ÂÊ  ÕœÌœ «··›Ÿ «··€ÊÏ «·„‰«”» ·ﬂ·„… Ã‰ÌÂ
    ''⁄·Ï Õ”» «·ﬁÌ„… «·„Œ’’… ·Â Ê«· Ï  ﬁ—« „‰ ﬁ«⁄œ… «·»Ì«‰« ..
    ''Ê«· Ï ÂÏ ﬁÌ„… «·›« Ê—…
    ''Ê‰Õ‰ ‰ﬁÿ⁄ „‰ ﬁÌ„… «·›« Ê—… «·„»·€ »«·Ã‰ÌÂ«  Ê‰Õ–› «·ﬁ—Ê‘
    ''..ÊÂ‰« ‰Õœœ „«–« ‰ﬂ »
    ''»ÃÊ«— –·ﬂ «·„»·€
    'Dim StrRS As ADODB.Recordset
    'Dim IntI As Integer
    'Dim StrArray 'this Array important in the Case of 103,104 or 1007 or 1010 and so on
    '
    'Set StrRS = New ADODB.Recordset
    'StrRS.Open "Options", cn, adOpenStatic, adLockReadOnly, adCmdTable
    '
    'StrArray = Array("03", "04", "05", "06", "07", "08", "09", "10")
    'For IntI = 0 To UBound(StrArray)
    '    If Right(CStr(DblThePoundValue), 2) = StrArray(IntI) Then
    '        If IsNull(StrRS("currency_unite_sum")) = False Then
    '            DefindPoundeType = StrRS("currency_unite_sum").Value
    '        Else
    '            DefindPoundeType = ""
    '        End If
    '        Exit Function
    '    End If
    'Next IntI
    '
    'If DblThePoundValue = 0 Then
    '     'ﬁÌ„… «·›« Ê—… √ﬁ· „‰ Ã‰ÌÂ ·–« ·‰ ‰ﬂ » ﬂ·„… Ã‰ÌÂ
    '    DefindPoundeType = ""
    'ElseIf DblThePoundValue = 1 Then  ' ›Ï Õ«·… «–« ﬂ«‰ «·„»·€ ÌÕ ÊÏ ⁄·Ï „ﬁÿ⁄ =1 Ã‰ÌÂ
    '    If IsNull(StrRS("currency_unite").Value) = False Then
    '        DefindPoundeType = StrRS("currency_unite").Value
    '    Else
    '        DefindPoundeType = ""
    '    End If
    '    'DefindPoundeType = "Ã‰ÌÂ"
    'ElseIf DblThePoundValue = 2 Then   '›Ï Õ«·… «–« ﬂ«‰ «·„»·€ ÌÕ ÊÏ ⁄·Ï „ﬁÿ⁄ =2 Ã‰ÌÂ
    '    If IsNull(StrRS("currency_unite_do")) = False Then
    '        DefindPoundeType = StrRS("currency_unite_do").Value
    '    Else
    '        DefindPoundeType = ""
    '    End If
    '    'DefindPoundeType = "Ã‰ÌÂ«‰"
    'ElseIf (3 <= DblThePoundValue And DblThePoundValue <= 10) Then
    '    If IsNull(StrRS("currency_unite_sum")) = False Then
    '        DefindPoundeType = StrRS("currency_unite_sum").Value
    '    Else
    '        DefindPoundeType = ""
    '    'DefindPoundeType = "Ã‰ÌÂ« "
    '    End If
    'Else '›Ï Ã„Ì⁄ «·Õ«·«  ›«‰ «··›Ÿ «··€ÊÏ ÌﬂÊ‰ Ã‰ÌÂ«
    '    If IsNull(StrRS("currency_unite_exp")) = False Then
    '        DefindPoundeType = StrRS("currency_unite_exp").Value
    '    Else
    '        DefindPoundeType = ""
    '    End If
    'End If
    'StrRS.Close
    'Set strts = Nothing

End Function

'
Function WriteNo(NO As String, Sex As Integer, Optional DecimalBracktes As Boolean = False, Optional DeciamlSymbol As String = "", Optional GroupingSymbol As String = "", Optional IntLang As Integer = 0, Optional DECIMAL_FOUND As Boolean = False _
, Optional currencycode As Integer, Optional IsEnglish As Integer = 0) As String


    If GET_DEFAULT_CURRENCY_INF(currencycode) = False Then
        FRMcurrency.show
        Exit Function
    End If

    If SystemOptions.UserInterface = EnglishInterface Or IntLang = 1 Then
        WriteNo = ConvertNumbersToWords(NO, currencycode)
        Exit Function
    End If

    Static FirstArray(9, 1) As String
    Static FirstArray1(2, 1)  As String
    Static SecondArray(9, 1) As String
    Static ThirdArray(9) As String

    ReDim Parts(4) As String
    ReDim PartStr(-1 To 3) As String

    Dim Length As Integer, i As Integer, TempLength As Integer
    Dim NoString As String, pos  As Integer
    Dim AfterPoint As String
    Dim Txt As String
    Dim StrSysDecSymbol As String
    Dim StrSysGroupSymbol As String
    Dim BolNegativeNumber As Boolean
    'sex=0 „–ﬂ—
    'sex= 1 „ƒ‰À

    'IntLang=0 Arabic
    'IntLang=1 English
    If IsEnglish Then GoTo EnglishN
    If SystemOptions.UserInterface = ArabicInterface Then
        FirstArray(1, 0) = "Ê«Õœ ": FirstArray(2, 0) = "«À‰«‰ ": FirstArray(3, 0) = "À·«À… "
        FirstArray(4, 0) = "√—»⁄… ": FirstArray(5, 0) = "Œ„”… ": FirstArray(6, 0) = "” … "
        FirstArray(7, 0) = "”»⁄… ": FirstArray(8, 0) = "À„«‰Ì… ": FirstArray(9, 0) = " ”⁄… "
    
        FirstArray(1, 1) = "Ê«Õœ… ": FirstArray(2, 1) = "«À‰ «‰ ": FirstArray(3, 1) = "À·«À "
        FirstArray(4, 1) = "√—»⁄ ": FirstArray(5, 1) = "Œ„” ": FirstArray(6, 1) = "”  "
        FirstArray(7, 1) = "”»⁄ ": FirstArray(8, 1) = "À„«‰ ": FirstArray(9, 1) = " ”⁄ "
                                                          
        FirstArray1(1, 0) = "√Õœ ": FirstArray1(2, 0) = "≈À‰« "
    
        FirstArray1(1, 1) = "≈ÕœÏ ": FirstArray1(2, 1) = "≈À‰ « "
                       
        SecondArray(1, 0) = "⁄‘—… ": SecondArray(2, 0) = "⁄‘—Ê‰ ": SecondArray(3, 0) = "À·«ÀÊ‰ "
        SecondArray(4, 0) = "√—»⁄Ê‰ ": SecondArray(5, 0) = "Œ„”Ê‰ ": SecondArray(6, 0) = "” Ê‰ "
        SecondArray(7, 0) = "”»⁄Ê‰ ": SecondArray(8, 0) = "À„«‰Ê‰ ": SecondArray(9, 0) = " ”⁄Ê‰ "
    
        SecondArray(1, 1) = "⁄‘—… ": SecondArray(2, 1) = "⁄‘—Ê‰ ": SecondArray(3, 1) = "À·«ÀÊ‰ "
        SecondArray(4, 1) = "√—»⁄Ê‰ ": SecondArray(5, 1) = "Œ„”Ê‰ ": SecondArray(6, 1) = "” Ê‰ "
        SecondArray(7, 1) = "”»⁄Ê‰ ": SecondArray(8, 1) = "À„«‰Ê‰ ": SecondArray(9, 1) = " ”⁄Ê‰ "
    
        ThirdArray(1) = "„«∆… ": ThirdArray(2) = "„«∆ «‰ ": ThirdArray(3) = "À·«À„«∆… "
        ThirdArray(4) = "√—»⁄„«∆… ": ThirdArray(5) = "Œ„”„«∆… ": ThirdArray(6) = "” „«∆… "
        ThirdArray(7) = "”»⁄„«∆… ": ThirdArray(8) = "À„«‰„«∆… ": ThirdArray(9) = " ”⁄„«∆… "
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
EnglishN:
        FirstArray(1, 0) = "One ": FirstArray(2, 0) = "Two ": FirstArray(3, 0) = "Three "
        FirstArray(4, 0) = "Four ": FirstArray(5, 0) = "Five ": FirstArray(6, 0) = "Six "
        FirstArray(7, 0) = "Seven ": FirstArray(8, 0) = "Eight ": FirstArray(9, 0) = "Nine "
    
        FirstArray(1, 1) = "One ": FirstArray(2, 1) = "Two ": FirstArray(3, 1) = "Three "
        FirstArray(4, 1) = "Four ": FirstArray(5, 1) = "Five ": FirstArray(6, 1) = "Six "
        FirstArray(7, 1) = "Seven ": FirstArray(8, 1) = "Eight ": FirstArray(9, 1) = "Nine "
                                                          
        FirstArray1(1, 0) = "√Õœ ": FirstArray1(2, 0) = "≈À‰« "
    
        FirstArray1(1, 1) = "≈ÕœÏ ": FirstArray1(2, 1) = "≈À‰ « "
                       
        SecondArray(1, 0) = "Ten ": SecondArray(2, 0) = "twenty ": SecondArray(3, 0) = "Thirty "
        SecondArray(4, 0) = "Forty ": SecondArray(5, 0) = "Fifty ": SecondArray(6, 0) = "Sixty "
        SecondArray(7, 0) = "Seventy ": SecondArray(8, 0) = "Eighty ": SecondArray(9, 0) = "Ninety "
    
        SecondArray(1, 1) = "Ten ": SecondArray(2, 1) = "Twenty ": SecondArray(3, 1) = "Thirty "
        SecondArray(4, 1) = "Forty ": SecondArray(5, 1) = "Fifty ": SecondArray(6, 1) = "Sixty "
        SecondArray(7, 1) = "Seventy ": SecondArray(8, 1) = "Eighty ": SecondArray(9, 1) = "Ninety "
    
        ThirdArray(1) = "One hundred": ThirdArray(2) = "two hundred ": ThirdArray(3) = "Three hundred "
        ThirdArray(4) = "Four hundred ": ThirdArray(5) = "Five hundred ": ThirdArray(6) = "Six hundred "
        ThirdArray(7) = "Seven hundred ": ThirdArray(8) = "Eight ": ThirdArray(9) = "Nine hundred "

    End If

    Txt = "": i = -1

    If val(NO) = 0 Then 'Â· «·⁄œœ «·„œŒ· ’›—
        WriteNo = "’›—"
        Exit Function
    End If

    '«Õ–› «·›—«€«  «·Ì„Ì‰Ì… Ê«·Ì”«—Ì… «·“«∆œ… ›Ì Õ«· ÊÃÊœÂ«
    NoString = Trim(NO)

    '----------------------
    'ÌÃ»  ÕœÌœ Â· «·—ﬁ„ ”«·» «„ „ÊÃ»
    If val(NO) > 0 Then
        BolNegativeNumber = False
    Else
        NO = Abs(val(NO))
        BolNegativeNumber = True
    End If

    '----------------------
    'ÌÃ» „⁄—›… ‰Ê⁄ «·›«’·… «·⁄‘—Ì…
    '«·„” Œœ„… ›Ï «·—ﬁ„ «·„—”·
    If DeciamlSymbol = "" Then
        StrSysDecSymbol = GetDeciamlSymbol
    Else
        StrSysDecSymbol = DeciamlSymbol
    End If

    '----------------------
    '·Ê «‰ «·—ﬁ„ «·„—”· ≈·Ï «·œ«·… »Â ⁄·«„… ⁄‘—Ì…
    '»œ·« „‰ «·⁄·«ﬁ… «·√› —«÷Ì… «·„Œ’’… ›Ï «·ÃÂ«“
    '›Ï Â–Â «·Õ«·… ·«»œ „‰  »œÌ· «·⁄·«„… «·⁄‘—Ì… «·„—”·…
    '›Ï «·—ﬁ„ ‰›”Â »«·⁄·«„… «·⁄‘—Ì… «·√› —«÷Ì… Ê–·ﬂ
    pos = InStr(NoString, ".")

    If pos > 0 Then
        If StrSysDecSymbol <> "." Then
            NoString = Replace(NoString, ".", StrSysDecSymbol, , , vbBinaryCompare)
        End If
    End If

    '----------------------
    'ÌÃ» „⁄—›… ‰Ê⁄ ⁄·«„…  ÃÌ„⁄ «·¬·«›
    If GroupingSymbol = "" Then
        StrSysGroupSymbol = GetGroupingSymbol
    Else
        StrSysGroupSymbol = GroupingSymbol
    End If

    '----------------------
    pos = InStr(NoString, ",")

    If pos > 0 Then
        If StrSysGroupSymbol <> "," Then
            NoString = Replace(NoString, ",", "", , , vbBinaryCompare)
        End If
    End If

    '----------------------
    pos = InStr(NoString, StrSysGroupSymbol)

    If pos > 0 Then
        If StrSysDecSymbol <> "." Then
            NoString = Replace(NoString, StrSysGroupSymbol, "", , , vbBinaryCompare)
        End If
    End If

    '----------------------
    If CheckTheSendNumber(NoString, StrSysDecSymbol, StrSysGroupSymbol) = False Then
        If IntLang = 0 Then
            WriteNo = "Œÿ√ ›Ï «·—ﬁ„ «·„—”· ...!!!"
        Else
            WriteNo = "Error in the Number...!!!"
        End If

        Exit Function
    End If

    '----------------------
    ' «Õ’· ⁄·Ï ÿÊ· ”·”·… «·⁄œœ
    Length = Len(NoString)
    '«Õ›Ÿ „ﬂ«‰ ÊÃÊœ «·›«’·… «·⁄‘—Ì…
    pos = InStr(NoString, StrSysDecSymbol)

    '«ﬁ”„ ”·”·… «·⁄œœ≈·Ï „«ﬁ»· «·›«’·… Ê„«»⁄œ «·›«’·…
    If pos > 0 Then
        AfterPoint = right$(NoString, Length - pos)
        NoString = left$(NoString, pos - 1)
        Length = Len(NoString)
    Else
        pos = InStr(NoString, ",")

        If pos > 0 Then
            AfterPoint = right$(NoString, Length - pos)
            NoString = left$(NoString, pos - 1)
            Length = Len(NoString)
        End If
    End If

    'Ã“¡ «·⁄œœ ≈·Ï ”·«”· Õ—›Ì… „ƒ·›… „‰ À·«À Œ«‰«  ⁄‘—Ì… √Ê √ﬁ·
    TempLength = Length
    Parts(0) = NoString

    Do While TempLength >= 3
        TempLength = TempLength - 3
        i = i + 1
        Parts(i) = right$(NoString, 3)
        NoString$ = left$(NoString, TempLength)
    Loop

    Parts(i + 1) = NoString

    '«” œ⁄ «· «»⁄ «·›—⁄Ì Ê«Õ›Ÿ «·‰ «∆Ã ›Ì «·„’›Ê›…
    For i = 0 To 3

        If Len(Parts(i)) > 0 Then
            PartStr(i) = GetNo(Parts(i), Sex, i, FirstArray(), FirstArray1(), SecondArray(), ThirdArray(), IntLang)
        Else
            Exit For
        End If

    Next

    '««Ã„⁄ «·ﬂ·„«  «·Ã“∆Ì… «·‰« Ã… ›Ì ⁄«—… Ê«Õœ…
    For i = 3 To 0 Step -1

        If Len(PartStr(i)) > 0 Then
            If Len(PartStr(i - 1)) > 0 Then
                Txt = Txt & " " & PartStr(i) & IIf(IntLang = 0, "Ê", "and")
            Else
                Txt = Txt & " " & PartStr(i) & " "
            End If
        End If

    Next i

    If DECIMAL_FOUND = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Txt = "  ›ﬁÿ " & Txt & " " & DEFALUT_CURRENCY
        Else
            Txt = " ONLY " & Txt & " " & DEFALUT_CURRENCYE
        End If
    End If

    If val(AfterPoint) > 0 Then
        Dim StrTemp As String
        StrTemp = GetAfterPoint(AfterPoint)
        DecimalBracktes = False

        If DecimalBracktes = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Txt = Txt & " Ê ) " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp & " ("
            Else
                Txt = Txt & " and (" & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp & ")"
            End If

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                Txt = Txt & " Ê " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp
            Else
                Txt = Txt & " and " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp
            End If
        End If
    End If

    If BolNegativeNumber = True Then
        If IntLang = 0 Then
            Txt = "”«·» " & Txt
        Else
            Txt = "negative " & Txt
        End If
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        'WriteNo = " ›ﬁÿ " & Txt & "" & Get_currency_txt & "  ·«€Ì— "
        WriteNo = Txt & "  ·«€Ì— "

    Else
        'WriteNo = Txt & "" & Get_currency_txt
        WriteNo = Txt

    End If

    If DECIMAL_FOUND = True Then
        WriteNo = Txt
        Exit Function
    End If
 
End Function

Public Function ConvertNumbersToWords(ByVal strInput As String, Optional currencycode As Integer) As String

    Dim strCleaned As String
    Dim nLoop As Integer
    Dim nLoop2 As Integer
    Dim strCents As String
    Dim strDollars As String
    Dim strConverted As String
    Dim strConvertedAll As String
    Dim strSection As String
    Dim strSectionValue As String
    Dim strSubSection As String
    Dim strNbrValue As String
    Dim strNbrValue2 As String

GET_DEFAULT_CURRENCY_INF currencycode
    ' Remove any characters not numeric or decimal
    For nLoop = 1 To Len(strInput)

        Select Case mId$(strInput, nLoop, 1)

            Case "0" To "9", "."
                strCleaned = strCleaned & mId$(strInput, nLoop, 1)
        End Select

    Next

    ' Check for cents
    nLoop = InStr(strCleaned, ".")

    ' Pad both with zeros
    If nLoop > 0 Then
        strCents = right$("00" & mId$(strCleaned, nLoop + 1), 3)
        strDollars = right$(String$(12, "0") & left$(strCleaned, nLoop - 1), 12)
    Else
        strDollars = right$(String$(12, "0") & strCleaned, 12)
        strCents = "00"
    End If

    ' Put back together
    strCleaned = strDollars & "." & strCents

    ' Start making words
    For nLoop = 1 To Len(strCleaned)

        ' Which section of the number are we on?
        Select Case nLoop

            Case 1
                strConverted = vbNullString
                strSectionValue = vbNullString
                strSection = "Billion "

            Case 4

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Million "

            Case 7

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Thousand "

            Case 10

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = vbNullString

            Case 14

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Cents"
        End Select
  
        If mId$(strCleaned, nLoop, 1) = "." Then
        Else

            For nLoop2 = 1 To 3

                ' Number value
                Select Case nLoop2

                    Case 1, 3

                        Select Case val(mId$(strCleaned, nLoop2 + nLoop - 1, 1))

                            Case 1: strNbrValue = "One"

                            Case 2: strNbrValue = "Two"

                            Case 3: strNbrValue = "Three"

                            Case 4: strNbrValue = "Four"

                            Case 5: strNbrValue = "Five"

                            Case 6: strNbrValue = "Six"

                            Case 7: strNbrValue = "Seven"

                            Case 8: strNbrValue = "Eight"

                            Case 9: strNbrValue = "Nine"

                            Case Else: strNbrValue = vbNullString
                        End Select
                             
                        Select Case nLoop2

                            Case 1

                                If strNbrValue <> vbNullString Then
                                    strSectionValue = strNbrValue & " Hundred"
                                End If

                            Case 3

                                If strNbrValue <> vbNullString Then
                                    If right$(strSectionValue, 3) = "Ten" Then

                                        Select Case strNbrValue

                                            Case "One":                           strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & "Eleven"

                                            Case "Two":                           strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & "Twelve"

                                            Case "Three":                         strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & "Thirteen"

                                            Case "Four":                          strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & "Fourteen"

                                            Case "Five":                          strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & "Fifteen"

                                            Case "Six", "Seven", "Eight", "Nine": strSectionValue = left$(strSectionValue, Len(strSectionValue) - 3) & strNbrValue & "teen"

                                            Case Else:                            strSectionValue = strSectionValue & " " & strNbrValue
                                        End Select

                                    Else
                                        strSectionValue = strSectionValue & " " & strNbrValue
                                    End If
                                End If

                        End Select

                    Case 2

                        Select Case val(mId$(strCleaned, nLoop2 + nLoop - 1, 1))

                            Case 1: strNbrValue2 = "Ten"

                            Case 2: strNbrValue2 = "Twenty"

                            Case 3: strNbrValue2 = "Thirty"

                            Case 4: strNbrValue2 = "Fourty"

                            Case 5: strNbrValue2 = "Fifty"

                            Case 6: strNbrValue2 = "Sixty"

                            Case 7: strNbrValue2 = "Seventy"

                            Case 8: strNbrValue2 = "Eighty"

                            Case 9: strNbrValue2 = "Ninety"

                            Case Else: strNbrValue2 = vbNullString
                        End Select

                        If strNbrValue2 <> vbNullString Then
                            strSectionValue = strSectionValue & " " & strNbrValue2
                        End If

                End Select

            Next

            nLoop = nLoop + 2
        End If

    Next

    ' Check for cents
    If strConvertedAll = "" Then strConvertedAll = "No "
    If Trim$(strSectionValue) = vbNullString Then
        If strConvertedAll = " One " Then
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE
        Else
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE
        End If

    Else

        If strConvertedAll = " One " Then
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE & "And" & strSectionValue & " " & DEFALUT_CURRENCY_DIVE
        Else
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE & " And" & strSectionValue & " " & DEFALUT_CURRENCY_DIVE
        End If
    End If

    ConvertNumbersToWords = strConvertedAll

End Function

Function GetNo(ns As String, Sex As Integer, Power As Integer, frst() As String, frst1() As String, scnd() As String, thrd() As String, Optional IntLang As Integer = 0) As String

    Dim Lngth As Integer, InvSex  As Integer
    ReDim Indx(3) As Integer
    ReDim TmpArray(2) As String
    Dim tms As String

    If Sex = 0 Then
        InvSex = 1
    Else
        InvSex = 0
    End If

    '«·Õ· „‰ √Ã· À·«À… √—ﬁ«„
    Lngth = Len(ns)
    '«·¬Õ«œ
    Indx(1) = val(mId$(ns, Lngth, 1))
    TmpArray(0) = frst(Indx(1), Sex)
    Lngth = Lngth - 1

    If Lngth > 0 Then
        '«·⁄‘—« 
        Indx(2) = val(mId$(ns, Lngth, 1))

        If TmpArray(0) <> "" Then
            TmpArray(1) = scnd(Indx(2), InvSex)
        Else
            TmpArray(1) = scnd(Indx(2), Sex)
        End If

        If (Indx(2) > 1) And (TmpArray(0) <> "") Then '«·⁄‘—«  „‰ 1 ≈·Ï  ”⁄…
            TmpArray(0) = TmpArray(0) & IIf(IntLang = 0, " Ê ", " and ")
        ElseIf (Indx(1) = 1) And (Indx(2) = 1) Then  '√Õœ ⁄‘—
            TmpArray(0) = frst1(1, Sex)
        ElseIf (Indx(1) = 2) And (Indx(2) = 1) Then ' «À‰« ⁄‘—
            TmpArray(0) = frst1(2, Sex)
        End If

        Lngth = Lngth - 1

        If Lngth > 0 Then
            '«·„∆« 
            Indx(3) = val(mId$(ns, Lngth, 1))
            TmpArray(2) = thrd(Indx(3))

            If (Indx(3) > 0) And ((TmpArray(0) <> "") Or (TmpArray(1) <> "")) Then
                TmpArray(2) = TmpArray(2) & IIf(IntLang = 0, " Ê ", " and ")
            End If

        Else
            GoTo last
        End If

    Else
        GoTo last
    End If

    '≈÷«›… ﬂ·„… «·„— »…(„∆…,√·›,...)Õ”» „— »… «·√—ﬁ«„
last:

    Select Case Power

        Case Is = -1
            tms = TmpArray(2) & TmpArray(0) & TmpArray(1)

            If (TmpArray(0) <> "") And (TmpArray(1) = "") And (TmpArray(2) = "") Then
                GetNo = tms & ""
            ElseIf (TmpArray(0) <> "") And (TmpArray(1) <> "") And (TmpArray(2) = "") Then
                GetNo = tms & ""
            ElseIf (TmpArray(0) <> "") And (TmpArray(1) <> "") And (TmpArray(2) <> "") Then
                GetNo = tms & ""
            End If

        Case Is = 0
            GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)

        Case Is = 1

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " √·› "
                GetNo = IIf(IntLang = 0, " √·› ", " Thousand ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " √·›«‰ "
                GetNo = IIf(IntLang = 0, " √·›«‰ ", " Two Thousand ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " ¬·«› "
                GetNo = IIf(IntLang = 0, TmpArray(0) & " ¬·«› ", TmpArray(0) & " Thousands")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then

                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " ¬·«› "
                If IntLang = 0 Then
                    GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " ¬·«› "
                ElseIf IntLang = 1 Then
                    GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " Thousands "
                End If

            ElseIf (Indx(1) = 0) And (Indx(2) = 0) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)
            Else
                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " √·› "
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " √·› ", " Thousand ")
            End If

        Case Is = 2

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·ÌÊ‰ "
                GetNo = IIf(IntLang = 0, " „·ÌÊ‰ ", " One Million ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·ÌÊ‰«‰ "
                GetNo = IIf(IntLang = 0, " „·ÌÊ‰«‰ ", " Two Millions ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " „·«ÌÌ‰ "
                GetNo = TmpArray(0) & IIf(IntLang = 0, " „·«ÌÌ‰ ", " Millions ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·«ÌÌ‰ ", " Millions ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 0) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)
            Else
                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " „·ÌÊ‰ "
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·ÌÊ‰ ", " Millions ")
            End If

        Case Is = 3

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·Ì«— "
                GetNo = IIf(IntLang = 0, " „·Ì«— ", " Milliard ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·Ì«—«‰ "
                GetNo = IIf(IntLang = 0, " „·Ì«—«‰ ", " Two Milliard ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " „·Ì«—«  "
                GetNo = TmpArray(0) & IIf(IntLang = 0, " „·Ì«—«  ", " Milliard ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·Ì«—«  ", " Milliard ")
            Else
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·Ì«— ", " Milliard ")
            End If

    End Select

End Function

Public Function headerdate(ConvertDate As Date) As String
    headerdate = Format(ConvertDate, "yyyy/m/d")
End Function

Public Function GetAfterPoint(AfPont As String) As String
    Dim StrTemp As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrTemp = DEFALUT_CURRENCY_DIV
    Else
        StrTemp = DEFALUT_CURRENCY_DIVE
    End If

    GoTo ll

    If Len(AfPont) = 1 Then
        StrTemp = "„‰ ⁄‘—…"
    ElseIf Len(AfPont) = 2 Then
        StrTemp = "„‰ „«∆…"
    ElseIf Len(AfPont) = 3 Then
        StrTemp = "„‰ «·›"
    ElseIf Len(AfPont) = 4 Then
        StrTemp = "„‰ ⁄‘—… √·«›"
    ElseIf Len(AfPont) = 5 Then
        StrTemp = "„‰ „«∆… «·›"
    ElseIf Len(AfPont) = 6 Then
        StrTemp = "„‰ „·ÌÊ‰"
    ElseIf Len(AfPont) = 7 Then
        StrTemp = "„‰ ⁄‘—… „·«Ì‰"
    ElseIf Len(AfPont) = 8 Then
        StrTemp = "„‰ „«∆… „·ÌÊ‰"
    ElseIf Len(AfPont) = 9 Then
        StrTemp = "„‰ „·Ì«—"
    ElseIf Len(AfPont) = 10 Then
        StrTemp = "„‰ ⁄‘—… „·Ì«—"
    ElseIf Len(AfPont) = 11 Then
        StrTemp = "„‰ „«∆… „·Ì«—"
    ElseIf Len(AfPont) = 12 Then
        StrTemp = "„‰  —Ì·ÌÊ‰"
    ElseIf Len(AfPont) = 13 Then
        StrTemp = "„‰ ⁄‘—…  —Ì·ÌÊ‰"
    ElseIf Len(AfPont) = 14 Then
        StrTemp = "„‰ „«∆…  —Ì·ÌÊ‰"
    Else
        StrTemp = "€Ì— „Õœœ"
    End If

ll:
    GetAfterPoint = StrTemp
End Function

Private Function GetDeciamlSymbol() As String
    Dim i As Single
    Dim StrTemp As String
    i = 1 / 2
    StrTemp = FormatNumber(i, , vbUseDefault, vbUseDefault, vbTrue)
    GetDeciamlSymbol = mId$(StrTemp, 2, 1)
End Function

Public Function GetGroupingSymbol() As String
    Dim i As Single
    Dim StrTemp As String
    i = 8819
    StrTemp = FormatNumber(i, , vbUseDefault, vbUseDefault, vbTrue)
    GetGroupingSymbol = mId$(StrTemp, 2, 1)
End Function

Private Function CheckTheSendNumber(StrNumber As String, _
                                    StrSysDecSymbol As String, _
                                    StrSysGroupSymbol As String) As Boolean

    Dim i As Integer
    Dim StrDigit As String

    For i = 1 To Len(StrNumber)
        StrDigit = mId$(StrNumber, i, 1)

        If InStr(1, "-0123456789" & StrSysDecSymbol & StrSysGroupSymbol, StrDigit, vbBinaryCompare) = 0 Then
            CheckTheSendNumber = False
            Exit Function
        End If

    Next i

    CheckTheSendNumber = True
End Function
