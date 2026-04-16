Attribute VB_Name = "Mod_Misc"
Option Explicit
#Const ProgVersion = "Demo"

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private gDirectPrintActive As Boolean
Private gDirectPrintToken As String
Private gDirectPrintKey As String
Private gDirectPrintLastKey As String
Private gDirectPrintLastTick As Long
Public Cn As New ADODB.Connection

Public user_id As Long

Public user_name As String

Public User_Password As String

Public bigUser As Boolean

Public StrCurName As String

Public VersionTest As Boolean

Public StrAppRegPath As String

Public SerialType As String
Public Declare Function InitCommonControls _
               Lib "comctl32.dll" () As Long

Public Enum ActionPostion
    SavePostion
    GetPostion
End Enum

Public Enum ActionSetting
    SaveMySetting
    GetMySetting
End Enum

Public Enum MySounds
    OpenscreenSnd
    AlreadyOpendSnd
    CloseScreenSnd
    ErrorSnd
End Enum

Public Enum PrintTarget
    WindowTarget
    PrinterTarget
End Enum



Public Enum ReportDirection
    ToWindow
    ToPrinter
End Enum
 Public Enum GridTransType
    InvoiceTransaction        '«·„»Ì⁄« 
    PurchaseTransaction        '«·‘—«¡
    Returntransaction    '„— Ã⁄ «·„‘ —Ì« 
    ShowPrice               '⁄—÷ √”⁄«—
    Maintenance                 '’Ì«‰…
    OpeningBalance       '—’Ìœ «›  «ÕÌ
    Template               '⁄—Ê÷ Ã«Â“…
    Destruction               '«·«Â·«ﬂ« 
    ReturnSalling        '„— Ã⁄ «·„»Ì⁄« 
    MoveItems  ' ÕÊÌ· «·»÷«⁄… „‰ „Œ“‰ ≈·Ï „Œ“‰
    InsertTemplate  '≈œ—«Ã ⁄—÷ Ã«Â“ ›Ì ⁄—Ê÷ «·√”⁄«—
    InsertTemplateToInvoice   '≈œ—«Ã ⁄—÷ Ã«Â“›Ì «·›« Ê—…
    PriceList            'ﬁ«∆„… «·√”⁄«—
    StockSettlement '  ”ÊÌ… «·„Œ“Ê‰
    InventoryOut '”‰œ ’—› „Œ“‰Ì
    INVENTORYIN '”‰œ «” ·«„  „Œ“‰Ì
    ProductionOrder  '«„— «‰ «Ã
        ProductionOrder1  '«„— «‰ «Ã
    RowMaterialIssue '”‰œ ’—› „Ê«œ Œ«„
    ProductionMaterialReciveVoucher '”‰œ «” ·«„ «‰ «Ã  «„
    NewGard       '  Ã—œ »‘ﬂ· ÃœÌœ
    purchaseorderrequest ' ÿ·» ⁄—÷ ”⁄— „‘ —Ì« 
    purchaseorder   '   ⁄—÷ ”⁄— „‘ —Ì« 
        internalissuerequesT    '   ÿ·» ’—› œ«Œ·Ì
        internalorder   '       ÿ·»«  œ«Œ·Ì…
        BookInventories ' ÕÃ“ »÷«⁄Â
purchaseOrderApproved ' «„— ‘—¡  ⁄„Ìœ
salespricelistRequest '  ÿ·»«  ⁄—Ê÷ «·«”⁄«—
salespricelist '    ⁄—Ê÷ «·«”⁄«—
SalesOrderRequest ' «Ê«„— «·»Ì⁄ «·„»œ∆Ì…
RowMaterialIssuesteps ' ”‰œ ’—› „—«Õ· «‰ «Ã
ProductionMaterialReciveVoucherStEPS ' ”‰œ «” ·«„ „—«Õ· « «Ã
ShipmentOrder 'ÿ·» ‘Õ‰
ShipmentRegisteration '  ”ÃÌ· »Ì«‰«  «·‘Õ‰
ShipmentRecieveVoucher ' ”‰œ «” ·«„ ‘Õ‰Â
RecervieProductionVoucherNew '”‰œ ÕÃ“ «”„‰ 
purchaserequest
ReturnDestruction
InvoiceTransactionCompose   '”‰œ ›« Ê—… „»Ì⁄«   Ã„Ì⁄Ï
End Enum


Public Function BeginDirectPrintGuard(ByVal PrintKey As String, _
                                      ByRef GuardToken As String, _
                                      Optional ByVal DuplicateWindowMs As Long = 2500) As Boolean
    Dim nowTick As Long
    Dim ageMs As Long

    nowTick = GetTickCount()
    ageMs = nowTick - gDirectPrintLastTick
    If ageMs < 0 Then ageMs = 0

    If gDirectPrintActive Then
        Debug.Print "DirectPrintGuard: blocked because another direct print is still active."
        Exit Function
    End If

    If LenB(gDirectPrintLastKey) > 0 Then
        If StrComp(gDirectPrintLastKey, PrintKey, vbTextCompare) = 0 Then
            If ageMs <= DuplicateWindowMs Then
                Debug.Print "DirectPrintGuard: blocked duplicate direct print. Age(ms)=" & CStr(ageMs)
                Exit Function
            End If
        End If
    End If

    GuardToken = "PRINT_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(nowTick)

    gDirectPrintActive = True
    gDirectPrintToken = GuardToken
    gDirectPrintKey = PrintKey

    BeginDirectPrintGuard = True
End Function

Public Sub EndDirectPrintGuard(ByVal GuardToken As String)
    Dim nowTick As Long

    If LenB(GuardToken) = 0 Then Exit Sub
    If StrComp(gDirectPrintToken, GuardToken, vbBinaryCompare) <> 0 Then Exit Sub

    nowTick = GetTickCount()

    gDirectPrintLastKey = gDirectPrintKey
    gDirectPrintLastTick = nowTick

    gDirectPrintActive = False
    gDirectPrintToken = ""
    gDirectPrintKey = ""
End Sub

Public Function NormalizePrintCopies(ByVal RequestedCopies As Long, _
                                     Optional ByVal MaxCopies As Long = 10) As Long
    NormalizePrintCopies = RequestedCopies

    If NormalizePrintCopies <= 0 Then
        NormalizePrintCopies = 1
    ElseIf NormalizePrintCopies > MaxCopies Then
        NormalizePrintCopies = MaxCopies
    End If
End Function

Public Function IsRemoteRedirectedPrinterName(ByVal printername As String) As Boolean
    Dim s As String

    s = UCase$(Trim$(printername))
    If LenB(s) = 0 Then Exit Function

    If InStr(s, "TSPLUS") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
    If InStr(s, "REDIRECT") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
    If InStr(s, "REMOTE DESKTOP") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
    If InStr(s, "EASY PRINT") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
    If InStr(s, "RDP") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
    If InStr(s, "(FROM ") > 0 Then IsRemoteRedirectedPrinterName = True: Exit Function
End Function

 

Public Sub SelectText(SelText As TextBox)
    On Error Resume Next
    SelText.SetFocus
    SelText.SelStart = 0
    SelText.SelLength = Len(SelText.Text)
End Sub
 
Public Sub clear_all(Frm As Form)
    
    Dim ctl As Control
    On Error Resume Next

    For Each ctl In Frm.Controls
        Debug.Print ctl.Name

        If TypeOf ctl Is ComboBox Then If ctl.Tag <> "not" Then ctl.ListIndex = -1
        If TypeOf ctl Is OptionButton Then If ctl.Tag <> "not" Then ctl.value = False
        If TypeOf ctl Is CheckBox Then If ctl.Tag <> "not" Then ctl.value = False
        If TypeOf ctl Is DataCombo Then If ctl.Tag <> "not" Then ctl.BoundText = ""
        
        If TypeOf ctl Is TextBox And ctl.Name <> "TxtModFlg" And ctl.Name <> "TxtModFlg1" And ctl.Name <> "TxtModFlg2" And ctl.Name <> "TxtModFlg3" And ctl.Name <> "TxtModFlg4" And ctl.Name <> "TxtModFlg5" And ctl.Name <> "TxtModFlg6" And ctl.Name <> "TxtModFlg7" And ctl.Name <> "TxtModFlg8" Then
            ctl.Text = ""
        Else
      '  X = 5
        End If

        '    If TypeOf Ctl Is TextBox And Ctl.name <> "not" Then Ctl.text = ""
        If TypeOf ctl Is DTPicker Then ctl.value = Date

        '    If TypeOf Ctl Is XPDatePicker30 Then Ctl.CurrentDate = ""
       If ctl.Tag = 1 Then
        ctl.Tag = 1
       End If
        
        If TypeOf ctl Is VSFlexGrid And ctl.Tag <> 1 Then
            If ctl.rows > 1 Then
                ctl.Clear 1, 1
                ctl.FixedRows = 1
                ctl.rows = ctl.FixedRows + 1
            End If
        End If

    Next

End Sub
 
Public Function checkfields(Frm As Form, _
                            Txt, _
                            Optional texts, _
                            Optional lbles) As Boolean
    On Error Resume Next
    Dim i As Integer

    If IsMissing(texts) Then

        For i = 0 To Frm.Txt.count - 1

            If InStr(1, Frm.Txt(i).Tag, "m") Then
                If Trim(Frm.Txt(i)) = "" Then
                    MsgBox "   √ﬂœ √‰ «·Õﬁ· ( " & Trim(Frm.lbl(i)) & " )€Ì— ›«—€ ", vbExclamation + vbDefaultButton1 + vbMsgBoxRight + vbMsgBoxRtlReading, "  ‰»ÌÂ "
                    On Error Resume Next
                    Frm.Txt(i).SetFocus
                    On Error GoTo 0
                    checkfields = False
                    Exit Function
                End If
            End If

        Next i

    Else

        For i = 0 To texts.count - 1

            If InStr(1, texts(i).Tag, "m") Then
                If Trim(texts(i)) = "" Then
                    MsgBox "   √ﬂœ √‰ «·Õﬁ· (" & lbles(i) & ") €Ì— ›«—€ ", vbExclamation + vbDefaultButton1 + vbMsgBoxRight + vbMsgBoxRtlReading, "  ‰»ÌÂ "
                    On Error Resume Next
                    texts(i).SetFocus
                    On Error GoTo 0
                    checkfields = False
                    Exit Function
                End If
            End If

        Next i

    End If

    checkfields = True
End Function

Public Function KeyAscii_Num(KeyAsc As Integer, _
                             Txt As String, _
                             Optional IntFilterType As Integer = 0) As Integer

    'IntFilterType=0 Readl Number
    'IntFilterType=1 Integer Number

    If KeyAsc = 8 Then
        KeyAscii_Num = KeyAsc
        Exit Function
    End If

    If IntFilterType = 0 Then
        If CBool(InStr(1, ".", Chr(KeyAsc))) And CBool(InStr(1, Txt, Chr(KeyAsc))) Then
            KeyAscii_Num = 0
            Exit Function
        ElseIf InStr(1, "0123456789.", Chr(KeyAsc)) = 0 Then
            KeyAscii_Num = 0
        Else
            KeyAscii_Num = KeyAsc
        End If

    ElseIf IntFilterType = 1 Then

        If InStr(1, "0123456789", Chr(KeyAsc)) = 0 Then
            KeyAscii_Num = 0
        Else
            KeyAscii_Num = KeyAsc
        End If
    End If

End Function

Public Sub Get_RetrunDate(Qty_Hour As Single, _
                          Out_Date As Date, _
                          Out_Time As Date, _
                          txtdate As TextBox, _
                          TxtTime As TextBox)
    Dim IntHour_No As Integer
    Dim IntDay_No As Integer
    Dim RetrunDate As Date
    Dim RetrunTime As Date
    Dim TempRetrunTime As Date
    Dim HaveDays As Boolean
    Dim InMorring As Boolean
    Qty_Hour = Qty_Hour * 24

    If Qty_Hour >= 24 Then
        HaveDays = True
    End If

    IntHour_No = Qty_Hour Mod 24

    If HaveDays = True Then
        IntDay_No = Int(Qty_Hour / 24)
    End If

    Debug.Print FormatDateTime(Out_Time, vbShortTime)

    If FormatDateTime(Out_Time, vbShortTime) < "12:00" Then
        InMorring = True
    Else
        InMorring = False
    End If

    'Calculate the The Retrun Day First
    If HaveDays = True Then
        RetrunDate = DateAdd("d", IntDay_No, Out_Date)
    Else
        RetrunDate = Out_Date
    End If

    If IntHour_No > 0 Then
        TempRetrunTime = DateAdd("h", IntHour_No, Out_Time)

        If InStr(1, CStr(TempRetrunTime), "31/12/1899", vbTextCompare) > 0 Then
            RetrunDate = DateAdd("d", 1, RetrunDate)
        End If

        RetrunTime = FormatDateTime(TempRetrunTime, vbLongTime)
    Else
        RetrunTime = Out_Time
    End If

    txtdate.Text = Format(RetrunDate, "yyyy/M/d")
    TxtTime.Text = FormatDateTime(RetrunTime, vbLongTime)
End Sub

Public Function WriteDate(Optional D_Date) As String
    Dim StrMSG As String
    Dim StrHijriDate As String
    Dim M_Date As Date

    If Not IsMissing(D_Date) Then
        M_Date = D_Date
    Else
        M_Date = Date
    End If

    StrMSG = ""

    Select Case Weekday(M_Date, vbSunday)

        Case vbSaturday
            StrMSG = StrMSG & " «·”»  "

        Case vbSunday
            StrMSG = StrMSG & " «·√Õœ "

        Case vbMonday
            StrMSG = StrMSG & " «·√À‰Ì‰ "

        Case vbTuesday
            StrMSG = StrMSG & " «·À·«À«¡ "

        Case vbWednesday
            StrMSG = StrMSG & " «·√—»⁄«¡ "

        Case vbThursday
            StrMSG = StrMSG & " «·Œ„Ì” "

        Case vbFriday
            StrMSG = StrMSG & " «·Ã„⁄… "
    End Select

    StrMSG = StrMSG & Format(M_Date, "yyyy/M/d", vbUseSystemDayOfWeek) & " „Ì·«œÌ… "
    StrMSG = StrMSG & "  " & Chr(13)
    VBA.Calendar = vbCalHijri
    StrHijriDate = " «·„Ê«›ﬁ "

    Select Case day(M_Date)

        Case 1
            StrHijriDate = StrHijriDate & "€‹‹—…"

        Case Else
            StrHijriDate = StrHijriDate & CStr(day(M_Date))
    End Select

    Select Case Month(M_Date)

        Case 1
            StrHijriDate = StrHijriDate & " „Õ—„ "

        Case 2
            StrHijriDate = StrHijriDate & " ’›— "

        Case 3
            StrHijriDate = StrHijriDate & " —»Ì⁄ √Ê· "

        Case 4
            StrHijriDate = StrHijriDate & "—»Ì⁄ À«‰Ï "

        Case 5
            StrHijriDate = StrHijriDate & " Ã„«œÏ √Ê·"

        Case 6
            StrHijriDate = StrHijriDate & " Ã„«œÏ À«‰Ï "

        Case 7
            StrHijriDate = StrHijriDate & " —Ã» "

        Case 8
            StrHijriDate = StrHijriDate & " ‘⁄»«‰ "

        Case 9
            StrHijriDate = StrHijriDate & " —„÷«‰ "

        Case 10
            StrHijriDate = StrHijriDate & " ‘ƒ«·"

        Case 11
            StrHijriDate = StrHijriDate & " –Ê «·ﬁ⁄œ… "

        Case 12
            StrHijriDate = StrHijriDate & " –Ê «·ÕÃ… "
    End Select

    StrHijriDate = StrHijriDate & " " & CStr(year(M_Date)) & " ÂÃ—Ì… "
    VBA.Calendar = vbCalGreg
    StrMSG = StrMSG & StrHijriDate
    WriteDate = StrMSG
End Function

Public Sub RunHelp()

    If Dir(App.path & "\Help\Help.exe") <> "" Then
        Shell App.path & "\Help\Help.exe", vbNormalFocus
    End If

End Sub

Public Sub CloseApplication()

    Dim i  As Integer
    On Error Resume Next

    Do While SystemOptions.BolUpdateTaskInProgress = True
        DoEvents
    Loop

    'Free the Hock on this Form
    'SetWindowLong MDIFrm.hwnd, GWL_WNDPROC, OrgProc
    'Free the Hock on the All application (Hock on the Msg box)
    'UnhookWindowsHookEx hHook
    'Unload all Forms
    'On Error GoTo ErrTrap
    i = 0

    Do

        If Forms(Forms.count - 1).Name <> "MDIFrmMain" Then
            Debug.Print Forms(Forms.count - 1).Name
            Unload Forms(Forms.count - 1)

            DoEvents
        End If

        'I = I + 1
    Loop While Forms.count > 1

    If Cn.State = adStateOpen Then
        Cn.Close
        Set Cn = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub FormPostion(Frm As Form, _
                       m_Pos As ActionPostion)
    
Exit Sub
    Dim StrSetting As String
    Dim ScreenSetting As String
    Dim VarSet As Variant

    If m_Pos = SavePostion Then
        SaveSetting StrAppRegPath, "FormsPostions\" & user_name & " \Resolution\" & (Screen.Width / Screen.TwipsPerPixelX), Frm.Name, Frm.left & "-" & Frm.top
    
    ElseIf m_Pos = GetPostion Then
        StrSetting = GetSetting(StrAppRegPath, "FormsPostions\" & user_name & " \Resolution\" & (Screen.Width / Screen.TwipsPerPixelX), Frm.Name, "")

        If StrSetting <> "" Then
            VarSet = Split(StrSetting, "-", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
                Frm.left = val(VarSet(0))
                Frm.top = val(VarSet(1))
            End If
        End If

        '    If ScreenSetting <> "" Then
        '        If ScreenSetting <> (Screen.Width / Screen.TwipsPerPixelX) Then
        '            CenterForm Frm
        '        End If
        '    End If
    End If

End Sub

Public Function Write_Qast(IntNo As Integer) As String
    Dim Temp As String

    If IntNo > 100 Then
        Exit Function
    End If

    Temp = Choose(IntNo, "«·√Ê·", "«·À«‰Ï", "«·À«·À", "«·—«»⁄", "«·Œ«„”", _
       "«·”«œ”", "«·”«»⁄", "«·À«„‰", "«· «”⁄", "«·⁄«‘—", "«·Õ«œÏ ⁄‘—", _
       "«·À«‰Ï ⁄‘—", "«·À«·À ⁄‘—", "«·—«»⁄ ⁄‘—", "«·Œ«„” ⁄‘—", "«·”«œ” ⁄‘—", _
       "«·”«»⁄ ⁄‘—", "«·À«„‰ ⁄‘—", "«· «”⁄ ⁄‘—", "«·⁄‘—Ì‰", "«·Õ«œÏ Ê«·⁄‘—Ì‰", _
       "«·À«‰Ï Ê«·⁄‘—Ì‰", "«·À«·À Ê«·⁄‘—Ì‰", "«·—«»⁄ Ê«·⁄‘—Ì‰", "«·Œ«„” Ê«·⁄‘—Ì‰", _
       "«·”«œ” Ê«·⁄‘—Ì‰", "«·”«»⁄ Ê«·⁄‘—Ì‰", "«·À«„‰ Ê«·⁄‘—Ì‰", "«· «”⁄ Ê«·⁄‘—Ì‰" _
       , "«·À·«ÀÌ‰", "«·Õ«œÏ Ê«·À·«ÀÌ‰", "«·À«‰Ï Ê«·À·«ÀÌ‰", "«·À«·À Ê«·À·«ÀÌ‰", "«·—«»⁄ Ê«·À·«ÀÌ‰", _
       "«·Œ«„” Ê«·À·«ÀÌ‰", "«·”«œ” Ê«·À·«ÀÌ‰", "«·”«»⁄ Ê«·À·«ÀÌ‰", "«·À«„‰ Ê«·À·«ÀÌ‰", "«· «”⁄ Ê«·À·«ÀÌ‰", _
       "«·√—»⁄Ì‰", "«·Õ«œÏ Ê«·√—»⁄Ì‰", "«·À«‰Ï Ê«·√—»⁄Ì‰", "«·À«·À Ê«·√—»⁄Ì‰", "«·—«»⁄ Ê«·√—»⁄Ì‰", _
       "«·Œ«„” Ê«·√—»⁄Ì‰", "«·”«œ” Ê«·√—»⁄Ì‰", "«·”«»⁄ Ê«·√—»⁄Ì‰", "«·À«„‰ Ê«·√—»⁄Ì‰", "«· «”⁄ Ê«·√—»⁄Ì‰", _
       "«·Œ„”Ì‰", "«·Õ«œÏ Ê«·Œ„”Ì‰", "«·À«‰Ï Ê«·Œ„”Ì‰", "«·À«·À Ê«·Œ„”Ì‰", "«·—«»⁄ Ê«·Œ„”Ì‰", _
       "«·Œ«„” Ê«·Œ„”Ì‰", "«·”«œ” Ê«·Œ„”Ì‰", "«·”«»⁄ Ê«·Œ„”Ì‰", "«·À«„‰ Ê«·Œ„”Ì‰", "«· «”⁄ Ê«·Œ„”Ì‰", _
       "«·” Ì‰", "«·Õ«œÏ Ê«·” Ì‰", "«·À«‰Ï Ê«·” Ì‰", "«·À«·À Ê«·” Ì‰", "«·—«»⁄ Ê«·” Ì‰", _
       "«·Œ«„” Ê«·” Ì‰", "«·”«œ” Ê«·” Ì‰", "«·”«»⁄ Ê«·” Ì‰", "«·À«„‰ Ê«·” Ì‰", "«· «”⁄ Ê«·” Ì‰" _
       , "«·”»⁄Ì‰", "«·Õ«œÏ Ê«·”»⁄Ì‰", "«·À«‰Ï Ê«·”»⁄Ì‰", "«·À«·À Ê«·”»⁄Ì‰", "«·—«»⁄ Ê«·”»⁄Ì‰", _
       "«·Œ«„” Ê«·”»⁄Ì‰", "«·”«œ” Ê«·”»⁄Ì‰", "«·”«»⁄ Ê«·”»⁄Ì‰", "«·À«„‰ Ê«·”»⁄Ì‰", "«· «”⁄ Ê«·”»⁄Ì‰", _
       "«·À„«‰Ì‰", "«·Õ«œÏ Ê«·À„«‰Ì‰", "«·À«‰Ï Ê«·À„«‰Ì‰", "«·À«·À Ê«·À„«‰Ì‰", "«·—«»⁄ Ê«·À„«‰Ì‰", _
       "«·Œ«„” Ê«·À„«‰Ì‰", "«·”«œ” Ê«·À„«‰Ì‰", "«·”«»⁄ Ê«·À„«‰Ì‰", "«·À«„‰ Ê«·À„«‰Ì‰", "«· «”⁄ Ê«·À„«‰Ì‰", _
       "«· ”⁄Ì‰", "«·Õ«œÏ Ê«· ”⁄Ì‰", "«·À«‰Ï Ê«· ”⁄Ì‰", "«·À«·À Ê«· ”⁄Ì‰", "«·—«»⁄ Ê«· ”⁄Ì‰", _
       "«·Œ«„” Ê«· ”⁄Ì‰", "«·”«œ” Ê«· ”⁄Ì‰", "«·”«»⁄ Ê«· ”⁄Ì‰", "«·À«„‰ Ê«· ”⁄Ì‰", "«· «”⁄ Ê«· ”⁄Ì‰", "«·„«∆…")
    Write_Qast = Temp
End Function

Public Sub MyPlaySound(MySnd As MySounds)
    'Select Case MySnd
    '    Case OpenscreenSnd
    '        If Dir(App.Path & "\Sound\ImpulseClickz.wav") <> "" Then
    '            PlaySound App.Path & "\Sound\ImpulseClickz.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
    '        End If
    '    Case AlreadyOpendSnd
    '        If Dir(App.Path & "\Sound\ImpulseNONO.wav") <> "" Then
    '            PlaySound App.Path & "\Sound\ImpulseNONO.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
    '        End If
    'End Select
End Sub

Public Function GetHijriDate(Optional ByVal G_Date, _
                             Optional ByName As Boolean = False) As String
    Dim Temp As String
    Dim IntXX As Integer

    If IsMissing(G_Date) Then
        G_Date = Date
    End If

    IntXX = Calendar
    Calendar = vbCalHijri

    If ByName = True Then
        Temp = ""
        Temp = day(G_Date)
        Temp = Temp & "" & MonthName(Month(G_Date))
        Temp = Temp & "" & year(G_Date)
        GetHijriDate = Temp
    Else
        GetHijriDate = CStr(G_Date)
    End If

    Calendar = IntXX
End Function

Public Sub Resize_Form(Frm As Form, _
                       Optional SizeType As FormSizeType = NoChangeInSize)

    
    If Frm.WindowState = vbNormal Then
        If SizeType = TransactionSize Then
            Frm.Height = 10000
            Frm.Width = 16000
        ElseIf SizeType = ReportSize Then
            Frm.Height = 9240
            Frm.Width = 11100
        End If

        Frm.top = (mdifrmmain.ScaleHeight - Frm.Height) / 2
        Frm.left = (mdifrmmain.ScaleWidth - Frm.Width) / 2
    End If



End Sub

Public Function checkApility(Frm As String, _
                             Optional BolShowMsg As Boolean = True) As Boolean

    Dim StrSQL As String
    Dim Msg As String
    Dim RsAllowEdit As ADODB.Recordset

    On Error GoTo ErrTrap
 
    'If user_id <> 1 And SystemOptions.usertype <> UserAdminAll Then
    If user_id <> 1 Then
        StrSQL = "Select * From ScreenJuncUser where  User_ID =" & user_id
        StrSQL = StrSQL + " and ScreenName='" & Frm & "'  order by CanShow desc"
        Set RsAllowEdit = New ADODB.Recordset
        RsAllowEdit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsAllowEdit.EOF Or RsAllowEdit.BOF) Then
            If RsAllowEdit("CanShow").value = True Or RsAllowEdit("CanAdd").value = True Then
                RsAllowEdit.Close
                checkApility = True
                Exit Function
            Else

                            If BolShowMsg = True Then
                                                    If SystemOptions.UserInterface = ArabicInterface Then
                                                                        Msg = "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…"
                                                    Else
                                                               Msg = "You are not authorized to Work  with this screen"
                                                    End If
                                MsgBox Msg, vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                            End If
                
                                checkApility = False
                                Exit Function
            End If

        Else

            If BolShowMsg = True Then
                'Msg = "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…"
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                                        Msg = "€Ì— „”„ÊÕ ·ﬂ »«· ⁄«„· „⁄ Â–Â «·‰«›–…"
                                                    Else
                                                               Msg = "You are not authorized to Work  with this screen"
                                                    End If
                MsgBox Msg, vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
            End If

            checkApility = False
            Exit Function
        End If

    Else
        checkApility = True
    End If

    Exit Function
ErrTrap:
End Function

Public Sub SetDtpickerDate(Dtp As DTPicker)
    Dtp.CalendarBackColor = &HC0FFFF
    Dtp.CalendarForeColor = &H80000012
    Dtp.CalendarTitleBackColor = &H404040
    Dtp.CalendarTitleForeColor = &HC0FFFF
    Dtp.CalendarTrailingForeColor = &H80000011

'    Dtp.Format = dtpCustom

    If SystemOptions.UserInterface = ArabicInterface Then
'       Dtp.CustomFormat = "yyyy/MM/dd"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
'        Dtp.CustomFormat = "d/M/yyyy"
    End If

    Dtp.value = Date

    If Dtp.CheckBox = True Then
        Dtp.value = Null
    End If

End Sub

Public Function Loaded(formname As String) As Boolean
    Dim i As Integer
    Loaded = False

    For i = 0 To Forms.count - 1

        If Forms(i).Name = formname Then
            Loaded = True
            Exit Function
        End If

        Debug.Print Forms(i).Name
    Next i

End Function

Public Function DisplayCurrency(DblValue As Double) As Currency
    DisplayCurrency = Format(DblValue, SystemOptions.SysDefCurrencyForamt)
End Function

