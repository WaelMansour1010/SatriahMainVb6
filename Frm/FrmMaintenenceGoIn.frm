VERSION 5.00
Begin VB.Form FrmMaintenenceGoIn 
   BackColor       =   &H00E2E9E9&
   Caption         =   " ﬁ—Ì— Ê„ «»⁄… «·’Ì«‰…"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   Icon            =   "FrmMaintenenceGoIn.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9180
End
Attribute VB_Name = "FrmMaintenenceGoIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Rs As ADODB.Recordset
'Dim TTP As clstooltip
'Dim TTD As clstooltipdemand
'Dim MaintenReport As ClsMaintananceReport
'Dim cSearchDcbo(4) As clsDCboSearch
'Public BolPrint As Boolean
'
'Private Sub CboMaintenanceType_Change()
''On Error GoTo ErrTrap
''If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
''    If CboMaintenanceType.ListIndex = 0 Then
''        TxtTransSerial.Enabled = False
''        lbl(9).Enabled = False
''        CmdSearch.Enabled = False
''        CmdSearchTrans.Enabled = False
''        CmdOpenTrans.Enabled = False
''    Else
''        TxtTransSerial.Enabled = True
''        lbl(9).Enabled = True
''        CmdSearch.Enabled = True
''        CmdSearchTrans.Enabled = True
''        CmdOpenTrans.Enabled = True
''    End If
''End If
''Exit Sub
'ErrTrap:
'End Sub
'Private Sub CboMaintenanceType_Click()
'CboMaintenanceType_Change
'End Sub
'
'Private Sub CmdAdd_Click()
''“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·
'
'Dim Msg As String
'Dim ItemCount As Integer
'Dim StrSerial As String
'Dim VarNum As Integer
'Dim Rs As ADODB.Recordset
'Dim StrSQL As String
'Dim LngFindRow As Long
'Dim LngRow As Long
'
'On Error GoTo ErrTrap
'If DCboItemsCode.text = "" Then
'    Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    DCboItemsCode.SetFocus
'    SendKeys "{F4}"
'    Exit Sub
'End If
'If DCboItemsName.text = "" Then
'    Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    DCboItemsName.SetFocus
'    SendKeys "{F4}"
'    Exit Sub
'End If
'
'If Val(TxtQuantity.text) = 0 Then
'    Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    TxtQuantity.SetFocus
'    Exit Sub
'End If
'If Me.TxtSerial.Enabled = True And Trim(Me.TxtSerial.text) = "" Then
'    Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·’‰›...!!"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    TxtSerial.SetFocus
'    Exit Sub
'End If
'If Me.CboMaintenanceType.ListIndex = -1 Then
'    Msg = "»—Ã«¡  ÕœÌœ ‰Ê⁄ «·’Ì«‰… ﬁ»· ≈÷«›… «·’‰› ..!!"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    CboMaintenanceType.SetFocus
'    Exit Sub
'End If
'If Val(Me.DcboEmpDes.BoundText) = 0 Then
'    Msg = "»—Ã«¡  ÕœÌœ ﬁ—«— «·„ÊŸ› ﬁ»· ≈÷«›… «·’‰› ..!!"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    DcboEmpDes.SetFocus
'    Exit Sub
'End If
'If Me.DBCboClientName.Enabled = True And Val(Me.DBCboClientName.BoundText) = 0 Then
'    Msg = "»—Ã«¡  ÕœÌœ  ÕœÌœ «·„Ê—œ «·–Ï ”Ê› Ì „  ÕÊÌ· «·’‰› ·Â ..!!"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    DBCboClientName.SetFocus
'    Exit Sub
'End If
'If Me.TxtCost.Enabled = True And Val(Me.TxtCost.text) = 0 Then
'    Msg = "ÌÃ» ≈œŒ«· ﬁÌ„… «· ﬂ·›… «Ê ›—ﬁ «·”⁄— ..!!"
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    TxtCost.SetFocus
'    Exit Sub
'End If
'If Me.TxtModFlg.text = "N" Then
'    LngTransID = 0
'ElseIf Me.TxtModFlg.text = "E" Then
'    LngTransID = Val(Me.XPTxtMaintanenceID.text)
'End If
'StrSQL = "SELECT QryManStockComplete.* "
'StrSQL = StrSQL + " FROM dbo.QryManStockComplete(" & LngTransID & ") QryManStockComplete"
'StrSQL = StrSQL + " Where ItemID=" & Me.DCboItemsCode.BoundText
'If Trim$(Me.TxtSerial.text) <> "" Then
'    StrSQL = StrSQL + " AND ItemSerial='" & Trim$(Me.TxtSerial.text) & "'"
'End If
'Set Rs = New ADODB.Recordset
'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'If (Rs.BOF Or Rs.EOF) Then
'    Msg = "Â–Â «·ﬁÿ⁄… €Ì— „ÊÃÊœ… ›Ï «·„Œ“‰ ›⁄·«,,,"
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    Rs.Close
'    Set Rs = Nothing
'    Exit Sub
'End If
'If Val(Me.TxtTicketNO.text) <> 0 Then
'    LngFindRow = Fg.FindRow(Val(Me.TxtTicketNO.text), Fg.FixedRows, Fg.ColIndex("TicketNO"), False, True)
'End If
'With Fg
'    If LngFindRow <= 0 Then
'        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
'           .Rows = .Rows + 1
'        End If
'        LngRow = .Rows - 1
'    Else
'        LngRow = LngFindRow
'    End If
'    .TextMatrix(LngRow, .ColIndex("Name")) = DCboItemsName.BoundText
'    .TextMatrix(LngRow, .ColIndex("Code")) = DCboItemsName.BoundText
'    .TextMatrix(LngRow, .ColIndex("Serial")) = TxtSerial.text
'    .TextMatrix(LngRow, .ColIndex("Count")) = TxtQuantity.text
'    .TextMatrix(LngRow, .ColIndex("TicketNO")) = TxtTicketNO.text
'    .TextMatrix(LngRow, .ColIndex("EmpNotes")) = Txt(1).text
'    If TxtSerial.Tag = "T" Then
'        .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexChecked
'    ElseIf TxtSerial.Tag = "F" Then
'        .Cell(flexcpChecked, LngRow, .ColIndex("HaveSerial")) = flexUnchecked
'    End If
'    .TextMatrix(LngRow, .ColIndex("EmpDes")) = Me.DcboEmpDes.BoundText
'    If Me.DBCboClientName.Enabled And Me.DBCboClientName.BoundText <> "" Then
'        .TextMatrix(LngRow, .ColIndex("SupManName")) = Me.DBCboClientName.text
'        .Cell(flexcpData, LngRow, .ColIndex("SupManName")) = Me.DBCboClientName.BoundText
'    End If
'    If Me.XPDtbGoOutDtae.Enabled = True Then
'        .TextMatrix(LngRow, .ColIndex("GoOutDate")) = Me.XPDtbGoOutDtae.Value
'    End If
'    .TextMatrix(LngRow, .ColIndex("Cost")) = Val(Me.TxtCost.text)
'
'    .AutoSize 0, .Cols - 1, False
'End With
'DCboItemsCode.BoundText = ""
'DCboItemsName.BoundText = ""
'Me.DcboEmpDes.BoundText = ""
'TxtSerial.text = ""
'Me.TxtTicketNO.text = ""
'Me.TxtQuantity.text = ""
'Me.TxtCost.text = 0
'Me.Txt(1).text = ""
''XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Cost"), FG.Rows - 1, FG.ColIndex("Cost"))
''-------------------------------------------
'
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub CmdHelp_Click()
'SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
'End Sub
'Private Sub CmdReplace_Click()
''Dim Msg As String
''Dim StrSQL As String
''Dim RsSerial As New ADODB.Recordset
''Dim RsTemp As New ADODB.Recordset
''On Error GoTo ErrTrap
''If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) = "" Then
''    Msg = "ÌÃ»  ÕœÌœ «·’‰› «·–Ì  —€» ›Ì «” »œ«·Â "
''    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''    Exit Sub
''End If
''If DBCboClientName.text = "" Then
''    Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·" & Chr(13)
''    Msg = Msg + "«·–Ì ﬁ«„ »‘—«¡ Â–Â «·ﬁÿ⁄…"
''    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''    DBCboClientName.SetFocus
''    SendKeys "{F4}"
''    Exit Sub
''End If
''If CboMaintenanceType.ListIndex = 1 Then
''    If TxtTransSerial.text = "" Then
''        Msg = Msg + "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ " & Chr(13)
''        Msg = Msg + "«· Ì  „ »Ì⁄ Â–« «·’‰› ›ÌÂ«"
''        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''        TxtTransSerial.SetFocus
''        Exit Sub
''    End If
''End If
'''«· √ﬂœ √‰ «·ﬁÿ⁄… ﬁœ  „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ… ›Ì Õ«·… «·’Ì«‰…  »⁄ «·÷„«‰
''If CboMaintenanceType.ListIndex = 1 Then
''    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" Then
''        If FG.Cell(flexcpChecked, FG.Row, FG.ColIndex("HaveSerial")) = flexChecked Then
''            If FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
''                StrSQL = "select * From QryGuarantee where Item_ID=" & _
''                FG.TextMatrix(FG.Row, FG.ColIndex("Code")) & _
''                " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
''                StrSQL = StrSQL + " AND Transaction_Serial ='" & Val(TxtTransSerial.text) & "'"
''                StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
''                RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
''                If RsSerial.EOF Or RsSerial.BOF Then
''                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
''                    Msg = Msg + FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
''                    Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
''                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
''                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''                    XPTab301.CurrTab = 0
''                    FG.Row = FG.Row
''                    FG.Col = FG.ColIndex("Name")
''                    FG.ShowCell FG.Row, FG.ColIndex("Name")
''                    FG.SetFocus
''                    Exit Sub
''                End If
''
''
''                If IsNull(RsSerial("guaranteeTime").Value) Then
''                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
''                    Msg = Msg + FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
''                    Msg = Msg + "·Ì” ·Â« ÷„«‰"
''                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''                    XPTab301.CurrTab = 0
''                    FG.Row = FG.Row
''                    FG.Col = FG.ColIndex("Name")
''                    FG.ShowCell FG.Row, FG.ColIndex("Name")
''                    FG.SetFocus
''                    Exit Sub
''                End If
''                If (DateDiff("d", XPDtbGoInDtae.Value, DateAdd("m", RsSerial("guaranteeTime").Value, RsSerial("Transaction_Date").Value))) < 0 Then
''                    Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
''                    Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").Value & Chr(13)
''                    Msg = Msg + "–«  «·”Ì—Ì«·  " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
''                    Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").Value, "yyyy/m/d") & Chr(13)
''                    Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").Value & Chr(13)
''                    Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").Value & "  ‘Â—" & Chr(13)
''                    Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"
''                    If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
''                        XPTab301.CurrTab = 0
''                        FG.Row = FG.Row
''                        FG.Col = FG.ColIndex("Name")
''                        FG.ShowCell FG.Row, FG.ColIndex("Name")
''                        FG.SetFocus
''                        Exit Sub
''                    End If
''                End If
''                RsSerial.Close
''            Else
''                Msg = "ÌÃ»  ÕœÌœ «·”Ì—Ì«· «·Œ«’ »«·ﬁÿ⁄… «· Ì  —€» ›Ì «” »œ«·Â«"
''                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''                Exit Sub
''            End If
''        Else
''            Msg = "Â–Â «·⁄„·Ì… Œ«’… »«·√’‰«› «· Ì   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
''            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
''            Exit Sub
''        End If
''    End If
''End If
''FG.Tag = FG.Row
'''With FrmReplace
'''    .TxtTransID.text = Me.TxtTransID.text
'''    .TxtTransSerial.text = Me.TxtTransSerial.text
'''    .XPTxtMaintanenceID.text = XPTxtMaintanenceID.text
'''    .DCboItemsName.BoundText = FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
'''    .Tag = FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Code"))
'''    .TxtItemSerial.text = FG.TextMatrix(FG.Row, FG.ColIndex("Serial"))
'''    .Show vbModal
'''End With
''Exit Sub
''ErrTrap:
'End Sub
'
'
'
'
'
'Private Sub CmdShowTransItems_Click()
'Dim Msg As String
'
'If Me.DCboStoreName.BoundText = "" Then
'    Msg = "ÌÃ» ≈Œ Ì«— «”„ «·„Œ“‰ √Ê·« ....!!"
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    Exit Sub
'End If
'Load FrmManChooseItems
'Set FrmManChooseItems.MyForm = Me
'FrmManChooseItems.ShowManStockItems Me.DCboStoreName.BoundText, Me.DCboStoreName.text
'FrmManChooseItems.Show
'End Sub
'
'Private Sub DcboEmpDes_Change()
'Dim IntValue As Integer
'IntValue = Val(Me.DcboEmpDes.BoundText)
''----------------------
'Me.lbl(6).Enabled = False
'Me.lbl(7).Enabled = False
'Me.lbl(9).Enabled = False
'Me.lbl(23).Enabled = False
'Me.DBCboClientName.Enabled = False
'Me.TxtCost.Enabled = False
'XPDtbGoOutDtae.Enabled = False
'Txt(1).Enabled = False
''---------------------
'If IntValue = 0 Then
'    Exit Sub
'ElseIf IntValue = 9 Or IntValue = 10 Then
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 11 Then
'    Me.lbl(6).Enabled = True
'    Me.DBCboClientName.Enabled = True
'    Me.lbl(7).Enabled = True
'    Me.XPDtbGoOutDtae.Enabled = True
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 12 Then
'    'NO Action
'    'Disable ALL
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 13 Then
'    Me.lbl(9).Enabled = True
'    Me.TxtCost.Enabled = True
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 14 Then
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 15 Then
'    '—œ ≈·Ï «·⁄„Ì· €Ì— ﬁ«œ— ⁄·Ï «·√’·«Õ
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 16 Then
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'ElseIf IntValue = 17 Or IntValue = 18 Then
'    Me.lbl(9).Enabled = True
'    Me.TxtCost.Enabled = True
'    Me.Txt(1).Enabled = True
'    Me.lbl(23).Enabled = True
'End If
'End Sub
'
'Private Sub DcboEmpDes_Click(Area As Integer)
'DcboEmpDes_Change
'End Sub
'
'
'Private Sub DCboItemsCode_Change()
'On Error GoTo ErrTrap
'Dim StrSQL As String
'Dim RsTemp As ADODB.Recordset
'If DCboItemsCode.BoundText <> "" Then
'    DCboItemsName.BoundText = DCboItemsCode.BoundText
'Else
'    Exit Sub
'End If
'StrSQL = "select * From TblItems where ItemID=" & DCboItemsCode.BoundText
'Set RsTemp = New ADODB.Recordset
'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Not (RsTemp.EOF Or RsTemp.BOF) Then
'    If RsTemp("HaveSerial").Value = True Then
''        TxtSerial.Enabled = True
''        TxtQuantity.Enabled = False
''        TxtQuantity.Text = "1"
'        TxtSerial.Tag = "T"
'    ElseIf RsTemp("HaveSerial").Value = False Then
''        TxtSerial.Enabled = False
''        TxtQuantity.Enabled = True
''        TxtQuantity.Text = ""
'        TxtSerial.Tag = "F"
'    End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub DCboItemsName_Change()
'On Error GoTo ErrTrap
'Dim StrSQL As String
'Dim RsTemp As ADODB.Recordset
'If DCboItemsName.BoundText <> "" Then
'    DCboItemsCode.BoundText = DCboItemsName.BoundText
'Else
'    Exit Sub
'End If
'StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
'Set RsTemp = New ADODB.Recordset
'RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Not (RsTemp.EOF Or RsTemp.BOF) Then
'    If RsTemp("HaveSerial").Value = True Then
'        TxtSerial.Enabled = True
'       TxtQuantity.Enabled = False
'        TxtQuantity.text = "1"
'    ElseIf RsTemp("HaveSerial").Value = False Then
'        TxtSerial.Enabled = False
'        TxtQuantity.Enabled = True
'        TxtQuantity.text = ""
'    End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Ele_Click(Index As Integer)
'On Error GoTo ErrTrap
'If Index = 1 Then
'    If Me.WindowState = vbNormal Then
'        Me.WindowState = vbMaximized
'    Else
'        Me.WindowState = vbNormal
'    End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub Ele_DblClick(Index As Integer)
' On Error GoTo ErrTrap
'Select Case Index
'    Case 7
'        If Me.WindowState = vbNormal Then
'            Me.WindowState = vbMaximized
'        Else
'            Me.WindowState = vbNormal
'        End If
'End Select
'Exit Sub
'ErrTrap:
'End Sub
'
'
'Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo ErrTrap
'Dim RsSerial As New ADODB.Recordset
'Dim RsTemp As ADODB.Recordset
'Dim Msg As String
'Dim StrSQL As String
'If XPDtbGoInDtae.Value = "" Then
'    Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ⁄„·Ì… «·’Ì«‰…" & Chr(13)
'    Msg = Msg + "ﬁ»· ≈œŒ«· »Ì«‰«  «·√’‰«›"
'    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    XPDtbGoInDtae.SetFocus
'    Exit Sub
'End If
'If Col = Fg.ColIndex("Name") Then
'    If Fg.TextMatrix(Row, Fg.ColIndex("Name")) <> "" Then
'        Fg.TextMatrix(Row, Fg.ColIndex("Code")) = Fg.TextMatrix(Row, Fg.ColIndex("Name"))
'        If IsNumeric(Fg.TextMatrix(Row, Fg.ColIndex("Code"))) Then
'            StrSQL = "select * From TblItems where ItemID=" & _
'            Fg.TextMatrix(Row, Fg.ColIndex("Code"))
'            Set RsTemp = New ADODB.Recordset
'            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'            If RsTemp.EOF Or RsTemp.BOF Then
'                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                Exit Sub
'            Else
'                If RsTemp("HaveSerial").Value = True Then
'                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = True
'                Else
'                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = False
'                End If
'            End If
'        Else
'            Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'    End If
'End If
'If Col = Fg.ColIndex("Code") Then
'    If Fg.TextMatrix(Row, Fg.ColIndex("Code")) <> "" Then
'        Fg.TextMatrix(Row, Fg.ColIndex("Name")) = Fg.TextMatrix(Row, Fg.ColIndex("Code"))
'        StrSQL = "select * From TblItems where ItemID=" & _
'        Fg.TextMatrix(Row, Fg.ColIndex("Code")) & ""
'        If IsNumeric(Fg.TextMatrix(Row, Fg.ColIndex("Code"))) Then
'            Set RsTemp = New ADODB.Recordset
'            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'            If RsTemp.EOF Or RsTemp.BOF Then
'                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                Exit Sub
'            Else
'                If RsTemp("HaveSerial").Value = True Then
'                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = True
'                Else
'                    Fg.TextMatrix(Row, Fg.ColIndex("HaveSerial")) = False
'                End If
'            End If
'        Else
'            Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'    End If
'End If
'If CboMaintenanceType.ListIndex = 1 Then
'    If Fg.TextMatrix(Row, Fg.ColIndex("Code")) <> "" Then
'        If Fg.Cell(flexcpChecked, Row, Fg.ColIndex("HaveSerial")) = flexChecked Then
'            If Fg.TextMatrix(Row, Fg.ColIndex("Serial")) <> "" Then
'                StrSQL = "select * From QryGuarantee where Item_ID=" & _
'                Fg.TextMatrix(Row, Fg.ColIndex("Code")) & _
'                " and ItemSerial='" & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & "'"
'                StrSQL = StrSQL + " AND Transaction_Serial='" & Val(TxtTransSerial.text) & "'"
'                StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
'                RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'                If RsSerial.EOF Or RsSerial.BOF Then
'                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
'                    Msg = Msg + Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
'                    Msg = Msg + "·„ Ì „ »Ì⁄Â« ›Ì «·›« Ê—… «·„Õœœ…" & Chr(13)
'                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ —ﬁ„ «·›« Ê—… Ê«”„ «·⁄„Ì·"
'                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'                    '»Ì«‰«  «·›« Ê—… «· Ì  „ »Ì⁄ «·ﬁÿ⁄Â ›ÌÂ«
'                    StrSQL = "select * From QryGuarantee where Item_ID=" & _
'                    Fg.TextMatrix(Row, Fg.ColIndex("Code")) & _
'                    " and ItemSerial='" & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & "'"
'                    Set RsTemp = New ADODB.Recordset
'                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
'                        Msg = "·ﬁœ  „ »Ì⁄ «·ﬁÿ⁄… : " & Fg.Cell(flexcpTextDisplay, Row, Fg.ColIndex("Name")) & Chr(13)
'                        Msg = Msg + "–«  «·”Ì—Ì«· : " & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
'                        Msg = Msg + "≈·Ï «·⁄„Ì· : " & RsTemp("CusName").Value & Chr(13)
'                        Msg = Msg + "›Ì «·›« Ê—… —ﬁ„ : " & RsTemp("Transaction_ID").Value
'                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    End If
'                    XPTab301.CurrTab = 0
'                    Fg.Row = Row
'                    Fg.Col = Fg.ColIndex("Name")
'                    Fg.ShowCell Row, Fg.ColIndex("Name")
'                    Fg.SetFocus
'                    Exit Sub
'                End If
'                If IsNull(RsSerial("guaranteeTime").Value) Then
'                    Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· " & Chr(13)
'                    Msg = Msg + Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
'                    Msg = Msg + "·Ì” ·Â« ÷„«‰"
'                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    XPTab301.CurrTab = 0
'                    Fg.Row = Row
'                    Fg.Col = Fg.ColIndex("Name")
'                    Fg.ShowCell Row, Fg.ColIndex("Name")
'                    Fg.SetFocus
'                    Exit Sub
'                End If
'                If (DateDiff("d", XPDtbGoInDtae.Value, DateAdd("m", RsSerial("guaranteeTime").Value, RsSerial("Transaction_Date").Value))) < 0 Then
'                    Msg = Msg + "«‰ Â  „œ… «·÷„«‰ «·Œ«’…" & Chr(13)
'                    Msg = Msg + "»«·ﬁÿ⁄…   " & RsSerial("ItemName").Value & Chr(13)
'                    Msg = Msg + "–«  «·”Ì—Ì«·  " & Fg.TextMatrix(Row, Fg.ColIndex("Serial")) & Chr(13)
'                    Msg = Msg + "›ﬁœ  „ »Ì⁄Â« » «—ÌŒ   " & Format(RsSerial("Transaction_Date").Value, "yyyy/m/d") & Chr(13)
'                    Msg = Msg + "›Ì «·›« Ê—… —ﬁ„  " & RsSerial("Transaction_ID").Value & Chr(13)
'                    Msg = Msg + "Êﬂ«‰  „œ… «·÷„«‰    " & RsSerial("guaranteeTime").Value & "  ‘Â—" & Chr(13)
'                    Msg = Msg + "Â·  —€» ›Ì ’Ì«‰ Â«  »⁄ «·÷„«‰ø"
'                    If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbNo Then
'                        XPTab301.CurrTab = 0
'                        Fg.Row = Row
'                        Fg.Col = Fg.ColIndex("Name")
'                        Fg.ShowCell Row, Fg.ColIndex("Name")
'                        Fg.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                RsSerial.Close
'            End If
'        End If
'    End If
'End If
'XPTxtSum.text = Fg.Aggregate(flexSTSum, 1, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'On Error GoTo ErrTrap
'If Col = Fg.ColIndex("HaveSerial") Then
'Cancel = True
'End If
'With Fg
'If .TextMatrix(Row, .ColIndex("MType")) <> "" Then
'    If .TextMatrix(Row, .ColIndex("MType")) = 2 Then
'        If Col = .ColIndex("Cost") Then
'            .TextMatrix(Row, .ColIndex("Cost")) = 0
'            Cancel = True
'        End If
'    End If
'End If
'If .TextMatrix(Row, .ColIndex("HaveSerial")) <> "" Then
'    If .TextMatrix(Row, .ColIndex("HaveSerial")) = False Then
'        If Col = .ColIndex("Serial") Then
'            Cancel = True
'        End If
'    End If
'End If
'End With
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'    FrmAddNewItem.DealingForm = Maintenance
'    FrmAddNewItem.Show vbModal
'End If
'End Sub
'Private Sub Fg_Click()
'On Error GoTo ErrTrap
''«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«·
'Dim StrSQL As String
'Dim RsReplace As ADODB.Recordset
'
'With Fg
'    If .Col = -1 Then Exit Sub
'    If .Row <= 0 Then Exit Sub
'    If .TextMatrix(.Row, .ColIndex("Name")) <> "" Then
'        Me.DCboItemsCode.BoundText = .TextMatrix(.Row, .ColIndex("Name"))
'        Me.DCboItemsName.BoundText = .TextMatrix(.Row, .ColIndex("Name"))
'        Me.DcboEmpDes.BoundText = Val(.TextMatrix(.Row, .ColIndex("EmpDes")))
'        Me.TxtCost.text = Val(.TextMatrix(.Row, .ColIndex("Cost")))
'        If Val(.Cell(flexcpData, .Row, .ColIndex("SupManName"))) = 0 Then
'            Me.DBCboClientName.BoundText = ""
'        Else
'            Me.DBCboClientName.BoundText = Val(.Cell(flexcpData, .Row, .ColIndex("SupManName")))
'        End If
'        If .Cell(flexcpChecked, .Row, .ColIndex("HaveSerial")) = flexChecked Then
'            Me.TxtSerial.Enabled = True
'            Me.TxtQuantity.Enabled = False
'            Me.TxtQuantity.text = 1
'            Me.TxtSerial.text = .TextMatrix(.Row, .ColIndex("Serial"))
'        Else
'            Me.TxtSerial.Enabled = False
'            Me.TxtQuantity.Enabled = True
'            Me.TxtQuantity.text = .TextMatrix(.Row, .ColIndex("Count"))
'            Me.TxtSerial.text = ""
'        End If
'
'        Me.TxtTicketNO.text = .TextMatrix(.Row, .ColIndex("TicketNO"))
'        Me.Txt(1).text = .TextMatrix(.Row, .ColIndex("EmpNotes"))
'        If Fg.TextMatrix(.Row, Fg.ColIndex("GoOutDate")) <> "" Then
'            XPDtbGoOutDtae.Value = Fg.TextMatrix(.Row, Fg.ColIndex("GoOutDate"))
'        End If
'    End If
'End With
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Form_Load()
'Dim StrSQL As String
'Dim BGround As New ClsBackGroundPic
'Dim RsItems As New ADODB.Recordset
'Dim StrList As String
'Dim Dcombos As ClsDataCombos
'Dim RsTemp As ADODB.Recordset
'
'On Error GoTo ErrTrap
''Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
''Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
''Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
''Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
''Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
''Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
''Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
''Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
''Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
''
''XPTab301.CurrTab = 0
''Me.Height = 8580
''Me.Width = 9700
''Resize_Form Me
'''AddTip
''SetDtpickerDate Me.XPDtbGoInDtae
''SetDtpickerDate XPDtbGoOutDtae
''Set Dcombos = New ClsDataCombos
''Dcombos.GetEmployees Me.DcboEmp
''Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
''Dcombos.GetStores Me.DCboStoreName
''Dcombos.GetUsers Me.DCboUserName
''Dcombos.GetManEmpDes DcboEmpDes
''With CboMaintenanceType
''    .AddItem "Œ«—Ã «·÷„«‰"
''    .AddItem "œ«Œ· «·÷„«‰"
''End With
''Fg.WallPaper = BGround.Picture
''
''Set cSearchDcbo(0) = New clsDCboSearch
''Set cSearchDcbo(0).Client = Me.DBCboClientName
''Set cSearchDcbo(1) = New clsDCboSearch
''Set cSearchDcbo(1).Client = Me.DcboEmp
''
''Set cSearchDcbo(2) = New clsDCboSearch
''Set cSearchDcbo(2).Client = Me.DCboStoreName
''LoadTBR
''Set Rs = New ADODB.Recordset
''Rs.Open "Select * From  TblMaintenece Where ManOperationTypeID=2", Cn, _
''adOpenStatic, adLockOptimistic, adCmdText
''
''StrSQL = "Select * From TblItems"
''RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''StrList = Fg.BuildComboList(RsItems, "ItemName", "ItemID")
''If StrList <> "" Then
''    Fg.ColComboList(Fg.ColIndex("Name")) = "|" & StrList
''End If
''StrList = Fg.BuildComboList(RsItems, "ItemCode", "ItemID")
''If StrList <> "" Then
''    Fg.ColComboList(Fg.ColIndex("Code")) = "|" & StrList
''End If
''Fg.ColComboList(Fg.ColIndex("MType")) = "#1;»«· ﬂ·›…|#2; »⁄ «·÷„«‰"
'''---------------------------------------------------------------------
''Set RsTemp = New ADODB.Recordset
''StrSQL = "SELECT SupDecID, SupDecName From dbo.TblManSupDecs"
''StrSQL = StrSQL + " Where (DecType = 2) ORDER BY SupDecID"
''RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''If Not (RsTemp.BOF Or RsTemp.EOF) Then
''    StrList = Fg.BuildComboList(RsTemp, "SupDecName", "SupDecID")
''End If
''If StrList <> "" Then
''    Fg.ColComboList(Fg.ColIndex("EmpDes")) = "|" & StrList
''End If
''
'''---------------------------------------------------------------------
''FillItemData
''Retrive
''Me.TxtModFlg.text = "R"
''Exit Sub
'ErrTrap:
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo ErrTrap
'Dim I As Integer
'If Rs.State = adStateOpen Then
'    If Not (Rs.EOF Or Rs.BOF) Then
'        If Rs.EditMode <> adEditNone Then
'            Rs.CancelUpdate
'        End If
'    End If
'    Rs.Close
'    Set Rs = Nothing
'End If
'For I = LBound(cSearchDcbo) To UBound(cSearchDcbo)
'    Set cSearchDcbo(I) = Nothing
'Next I
'Set MaintenReport = Nothing
'Set TTP = Nothing
'Set TTD = Nothing
'Exit Sub
'ErrTrap:
'
'
'
'
'
'End Sub
'
'
'
'Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'With Button
'    Select Case .Key
'        Case "RemoveRow"
'            If Fg.Rows > 1 Then
'                If Fg.Rows = 2 Then
'                    Me.Fg.Clear flexClearScrollable, flexClearEverything
'                Else
'                    If Me.Fg.Rows > 1 Then
'                        If Me.Fg.Row <> Me.Fg.FixedRows - 1 Then
'                            Me.Fg.RemoveItem (Me.Fg.Row)
'                        End If
'                    End If
'                End If
'            End If
'    End Select
'End With
'End Sub
'
'Private Sub TxtCost_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, TxtCost.text, 0)
'End Sub
'
'
'Private Sub TxtModFlg_Change()
'
'Select Case Me.TxtModFlg.text
'    Case "R"
'        Me.Caption = " ﬁ—Ì— Ê„ «»⁄… «·’Ì«‰…"
'        Me.Cmd(2).Enabled = False
'        Me.Cmd(3).Enabled = False
'
'        Me.Cmd(0).Enabled = True
'        Me.Cmd(1).Enabled = True
'        Me.Cmd(4).Enabled = True
'        Me.Cmd(5).Enabled = True
'        Me.Cmd(7).Enabled = True
'
'        Me.XPBtnMove(0).Enabled = True
'        Me.XPBtnMove(1).Enabled = True
'        Me.XPBtnMove(2).Enabled = True
'        Me.XPBtnMove(3).Enabled = True
'
'        XPDtbGoInDtae.Enabled = False
'        XPDtbGoOutDtae.Enabled = False
'        DBCboClientName.Locked = True
'
'        Fg.Editable = flexEDNone
'        If Rs.RecordCount < 1 Then
'            Me.XPBtnMove(0).Enabled = False
'            Me.XPBtnMove(1).Enabled = False
'            Me.XPBtnMove(2).Enabled = False
'            Me.XPBtnMove(3).Enabled = False
'            Me.Cmd(1).Enabled = False
'            Me.Cmd(4).Enabled = False
'            Me.Cmd(5).Enabled = False
'            Me.Cmd(7).Enabled = False
'        End If
'        CboMaintenanceType.Locked = True
'        Ele(5).Enabled = False
'        Me.DcboEmp.Locked = True
'        Me.DCboStoreName.Locked = True
'    Case "N"
'        Me.Caption = " ﬁ—Ì— Ê„ «»⁄… «·’Ì«‰…( ÃœÌœ )"
'        Me.Cmd(2).Enabled = True
'        Me.Cmd(3).Enabled = True
'
'        Me.Cmd(0).Enabled = False
'        Me.Cmd(1).Enabled = False
'        Me.Cmd(4).Enabled = False
'        Me.Cmd(5).Enabled = False
'        Me.Cmd(7).Enabled = False
'
'        Me.XPBtnMove(0).Enabled = False
'        Me.XPBtnMove(1).Enabled = False
'        Me.XPBtnMove(2).Enabled = False
'        Me.XPBtnMove(3).Enabled = False
'
'        XPDtbGoInDtae.Enabled = True
'        XPDtbGoOutDtae.Enabled = True
'        DBCboClientName.Locked = False
'
'        Fg.Enabled = True
'        Fg.Rows = Fg.FixedRows
'        Fg.Rows = 2
'        Me.DBCboClientName.Locked = False
'        Fg.Editable = flexEDNone
'        XPDtbGoInDtae.Value = Date '
'        XPDtbGoOutDtae.Value = Date
'        CboMaintenanceType.Locked = False
'        CboMaintenanceType.ListIndex = 0
'
'        CboMaintenanceType_Change
'        Ele(5).Enabled = True
'        Me.DcboEmp.Locked = False
'        Me.DCboStoreName.Locked = False
'    Case "E"
'        Me.Caption = " ﬁ—Ì— Ê„ «»⁄… «·’Ì«‰…(  ⁄œÌ· )"
'        Me.Cmd(2).Enabled = True
'        Me.Cmd(3).Enabled = True
'
'        Me.Cmd(0).Enabled = False
'        Me.Cmd(1).Enabled = False
'        Me.Cmd(4).Enabled = False
'        Me.Cmd(5).Enabled = False
'        Me.Cmd(7).Enabled = False
'
'
'        Me.XPBtnMove(0).Enabled = False
'        Me.XPBtnMove(1).Enabled = False
'        Me.XPBtnMove(2).Enabled = False
'        Me.XPBtnMove(3).Enabled = False
'
'        XPDtbGoInDtae.Enabled = True
'        XPDtbGoOutDtae.Enabled = True
'        DBCboClientName.Locked = False
'        XPBtnNewClients.Enabled = True
'
'
'        Fg.Enabled = True
'        Me.DBCboClientName.Locked = False
'        CboMaintenanceType.Locked = False
'        Fg.Editable = flexEDNone
'        DBCboClientName_Change
'        CboMaintenanceType_Change
'        Ele(5).Enabled = True
'        Me.DcboEmp.Locked = False
'        Me.DCboStoreName.Locked = False
'End Select
'Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub TxtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ErrTrap
'If KeyCode = vbKeyReturn Then
'    CmdAdd_Click
'End If
'Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.text, 1)
'End Sub
'
'
'
'
'Private Sub TxtTransSerial_Change()
''Dim StrTemp As String
''If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
''    If Trim(Me.TxtTransSerial.text) = "" Then
''        Me.TxtTransID.text = ""
''    Else
''        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2)
''        If Trim$(Me.TxtTransID.text) <> StrTemp Then
''            Me.TxtTransID.text = StrTemp
''        End If
''    End If
''End If
'End Sub
'
'Private Sub XPBtnAdd_Click()
'On Error GoTo ErrTrap
'If Fg.TextMatrix(Fg.Rows - 1, Fg.ColIndex("Code")) <> "" Then
'    Fg.Rows = Fg.Rows + 1
'    Fg.Row = Fg.Rows - 1
'    Fg.Col = Fg.ColIndex("Code")
'    Fg.ShowCell Fg.Rows - 1, Fg.ColIndex("Code")
'    Fg.SetFocus
'End If
'Exit Sub
'ErrTrap:
'End Sub
'
'
'Private Sub TxtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim Rs As ADODB.Recordset
'Dim StrSQL As String
'Dim Msg As String
'
'If KeyCode = vbKeyReturn Then
'    If Val(Me.DCboStoreName.BoundText) = 0 Then
'        ShowNotifyToolTip Me.DCboStoreName, "«”„ «·„Œ“‰", "ÌÃ» ≈Œ Ì«— «”„ «·„Œ“‰"
'    End If
'    If Val(Me.TxtTicketNO.text) <> 0 Then
'        StrSQL = "SELECT     QTY, ItemID, ItemCode, ItemName, HaveSerial," & _
'        "ItemSerial, TicketNO, StoreID, StoreName " & _
'        "FROM dbo.QryManStockComplete(0) QryManStockComplete"
'        StrSQL = StrSQL + " Where TicketNO=" & Val(Me.TxtTicketNO.text) & ""
'        StrSQL = StrSQL + " AND StoreID=" & Val(Me.DCboStoreName.BoundText) & ""
'        Set Rs = New ADODB.Recordset
'        Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (Rs.BOF Or Rs.EOF) Then
'            Me.DCboItemsCode.BoundText = IIf(IsNull(Rs("ItemID").Value), "", Rs("ItemID").Value)
'            Me.DCboItemsName.BoundText = IIf(IsNull(Rs("ItemID").Value), "", Rs("ItemID").Value)
'            Me.TxtQuantity.text = IIf(IsNull(Rs("QTY").Value), "", Rs("QTY").Value)
'            Me.TxtSerial.text = IIf(IsNull(Rs("ItemSerial").Value), "", Rs("ItemSerial").Value)
'        Else
'            Msg = "—ﬁ„ «· ﬂ  «·–Ï «œŒ· Â €Ì— ’ÕÌÕ "
'            Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ —ﬁ„ «· ﬂ "
'            Msg = Msg & Chr(13) & "«Ê «·„Œ“‰ «·„Õœœ"
'            ShowNotifyToolTip Me.TxtTicketNO, "—ﬁ„ «· ﬂ  €Ì— ’ÕÌÕ", Msg
'        End If
'        Rs.Close
'        Set Rs = Nothing
'    End If
'End If
'End Sub
'Private Sub ShowNotifyToolTip(XControl As Control, StrTitle As String, StrText As String)
'Set TTD = New clstooltipdemand
'Set TTD.m_From = Me
'TTD.Style = TTBalloon
'TTD.Icon = TTIconError
'TTD.Centered = True
'TTD.RightToLeft = True
'TTD.CreateToolTip XControl.hwnd
'TTD.DelayTime = 250
'TTD.VisibleTime = 5000
'TTD.Title = StrTitle
'TTD.TipText = StrText
'TTD.PopupOnDemand = True
'TTD.Show (XControl.Width / Screen.TwipsPerPixelY), _
'(XControl.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
'
'End Sub
'
'Private Sub TxtTicketNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = vbKeyReturn Then
'    KeyAscii = 0
'End If
'End Sub
'
'Private Sub XPBtnMove_Click(Index As Integer)
'On Error GoTo ErrTrap
'Select Case Index
'    Case 0
'        If Not (Rs.EOF Or Rs.BOF) Then
'            Rs.MovePrevious
'            If Rs.BOF Then Rs.MoveFirst
'        End If
'    Case 1
'        If Not (Rs.EOF Or Rs.BOF) Then
'            Rs.MoveFirst
'        End If
'    Case 2
'        If Not (Rs.EOF Or Rs.BOF) Then
'            Rs.MoveLast
'        End If
'    Case 3
'        If Not (Rs.EOF Or Rs.BOF) Then
'            Rs.MoveNext
'            If Rs.EOF Then Rs.MoveLast
'        End If
'End Select
'Retrive
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Cmd_Click(Index As Integer)
'Dim StrSQL As String
'Dim RsTemp As ADODB.Recordset
'Dim AskOption As Boolean
'Dim intDef As Integer
'BolPrint = True
'Dim Msg As String
'On Error GoTo ErrTrap
'Select Case Index
'    Case 0
'        clear_all Me
'        TxtModFlg.text = "N"
'        Me.DCboUserName.BoundText = User_ID
'        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
'        DBCboClientName.BoundText = intDef
'        XPTxtMaintanenceID.text = CStr(new_id("TblMaintenece", "MaintananceID", "", True))
'        XPTab301.CurrTab = 0
'        Fg.SetFocus
'        Fg.Col = Fg.ColIndex("Code")
'        Fg.Row = Fg.Rows - 1
'    Case 1
'        '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
'        StrSQL = "select * From  Transactions where MaintenanceID=" & Val(Rs("MaintananceID").Value)
'        Set RsTemp = New ADODB.Recordset
'        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RsTemp.EOF Or RsTemp.BOF) Then
'            Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« «Â«"
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            Exit Sub
'        End If
'        TxtModFlg.text = "E"
'        Me.DCboUserName.BoundText = User_ID
'    Case 2
'        SaveData
'    Case 3
'       Call Undo
'    Case 4
'       Del_TransAction
'    Case 5
'        Load FrmMaintanenceSearch
'        FrmMaintanenceSearch.SearchType = 2
'        FrmMaintanenceSearch.Show vbModal
'    Case 7
'        AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)
'        If AskOption = False Then
'            FrmPrintOptions.Show vbModal
'        End If
'        If BolPrint = False Then
'            Exit Sub
'        End If
'        PrintingData
'    Case 6
'        Unload Me
'End Select
'Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub XPBtnRemove_Click()
'On Error GoTo ErrTrap
'If Fg.Rows = 2 Then
'    Fg.Clear flexClearScrollable, flexClearEverything
'Else
'    If Fg.Rows > 1 Then
'        If Fg.Row <> Fg.FixedRows - 1 Then
'            Fg.RemoveItem (Fg.Row)
'        End If
'    End If
'End If
'XPTxtSum.text = Fg.Aggregate(flexSTSum, 1, Fg.ColIndex("Cost"), Fg.Rows - 1, Fg.ColIndex("Cost"))
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub SaveData()
'Dim RsNotes As New ADODB.Recordset
'Dim RsDetails As New ADODB.Recordset
'Dim RsSerial As New ADODB.Recordset
'Dim RsCheckSerial As New ADODB.Recordset
'Dim RsTemp As ADODB.Recordset
'Dim RsReplace As ADODB.Recordset
'Dim RsReplaceDetails As ADODB.Recordset
'Dim StrSQL As String
'Dim RowNum As Integer
'Dim ReplaceID As Integer
'Dim Msg As String
'Dim BeginTrans As Boolean
'
'On Error GoTo ErrTrap
'If Me.TxtModFlg.text <> "R" Then
'     If CboMaintenanceType.ListIndex = -1 Then
'        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·’Ì«‰…"
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        CboMaintenanceType.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
'    If Me.DcboEmp.BoundText = "" Then
'        Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„ÊŸ›...!!!"
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        DcboEmp.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
'    If CboMaintenanceType.ListIndex = 1 Then
'
'    End If
'    If Me.DCboStoreName.BoundText = "" Then
'        Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰....!!! " & Chr(13)
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        DCboStoreName.SetFocus
'        SendKeys "{F4}"
'        Exit Sub
'    End If
'    If ItemsInGrid(Me.Fg, Fg.ColIndex("Name")) = -1 Then
'        Msg = "ÌÃ»  ”ÃÌ· «Ï √’‰«› ›Ï «·Õ—ﬂ…..!! " & Chr(13)
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Exit Sub
'    End If
'    If Me.TxtModFlg.text = "N" Then
'        Me.XPTxtMaintanenceID.text = CStr(new_id("TblMaintenece", "MaintananceID", "", True))
'        Rs.AddNew
'    ElseIf Me.TxtModFlg.text = "E" Then
'       StrSQL = "delete From TblMainteneceDetails where MaintananceID=" & Val(Rs("MaintananceID").Value)
'       Cn.Execute StrSQL, , adExecuteNoRecords
'       StrSQL = "delete From MaintenanceJuncTransaction where MaintananceID=" & Val(Rs("MaintananceID").Value)
'       Cn.Execute StrSQL, , adExecuteNoRecords
'       StrSQL = "delete From Transactions where MaintenanceID=" & Val(Rs("MaintananceID").Value)
'       Cn.Execute StrSQL, , adExecuteNoRecords
'    End If
'    RsDetails.Open "[TblMainteneceDetails]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    Cn.BeginTrans
'    BeginTrans = True
'    Rs("MaintananceID").Value = Val(XPTxtMaintanenceID.text)
'    Rs("CusID").Value = Null ' Me.DBCboClientName.BoundText ' IIf(DBCboClientName.BoundText = "", "", DBCboClientName.BoundText)
'    Rs("DateGoIN").Value = XPDtbGoInDtae.Value
'    Rs("DateGoOUT").Value = Null
'    Rs("GoOut").Value = 0
'    Rs("EmpID").Value = Me.DcboEmp.BoundText
'    Rs("StoreID").Value = Me.DCboStoreName.BoundText
'    Rs("UserID").Value = User_ID
'    If CboMaintenanceType.ListIndex = -1 Then
'        Rs("MType").Value = 0
'    Else
'        Rs("MType").Value = Val(CboMaintenanceType.ListIndex)
'    End If
'    Rs("Transaction_ID").Value = Null
'    Rs("ManOperationTypeID").Value = 2
'    Rs.update
'    For RowNum = 1 To Fg.Rows - 1
'        RsDetails.AddNew
'        RsDetails("MaintananceID").Value = Val(XPTxtMaintanenceID.text)
'        RsDetails("ItemID").Value = IIf(IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))), "", Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))))
'        If Not Fg.TextMatrix(RowNum, Fg.ColIndex("Name")) = "" Then
'            StrSQL = "select * From TblItems where ItemID=" & Fg.TextMatrix(RowNum, Fg.ColIndex("Name"))
'            RsCheckSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'            If Not (RsCheckSerial.EOF Or RsCheckSerial.BOF) Then
'                If RsCheckSerial("HaveSerial").Value = True Then
'                    RsDetails("ItemSerial").Value = IIf(IsNull(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))), "", Trim(Fg.TextMatrix(RowNum, Fg.ColIndex("Serial"))))
'                End If
'            End If
'            RsCheckSerial.Close
'        End If
'        RsDetails("Quantity").Value = Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count")))
'        RsDetails("TicketNO").Value = Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("TicketNO")))
'        RsDetails("CustomerNotes").Value = Null
'        RsDetails("EmpNotes").Value = Trim$(Fg.TextMatrix(RowNum, Fg.ColIndex("EmpNotes")))
'        RsDetails("SupDeci").Value = Val(Fg.TextMatrix(RowNum, Fg.ColIndex("EmpDes")))
'        If Val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("SupManName"))) = 0 Then
'            RsDetails("SupID").Value = Null
'        Else
'            RsDetails("SupID").Value = Val(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("SupManName")))
'        End If
'        RsDetails("Cost").Value = Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Cost")))
'        If IsDate(Fg.TextMatrix(RowNum, Fg.ColIndex("GoOutDate"))) Then
'            RsDetails("RetrunDate").Value = (Fg.TextMatrix(RowNum, Fg.ColIndex("GoOutDate")))
'        Else
'            RsDetails("RetrunDate").Value = Null
'        End If
'        RsDetails.update
'    Next RowNum
'
'
'CompleteSaving:
'    Cn.CommitTrans
'    BeginTrans = False
'    XPTxtCurrent.Caption = Rs.AbsolutePosition
'    XPTxtCount.Caption = Rs.RecordCount
'    Select Case Me.TxtModFlg.text
'        Case "N"
'            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
'            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
'            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
'            Cmd_Click (0)
'            Exit Sub
'            End If
'        Case "E"
'            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    End Select
'    TxtModFlg.text = "R"
'End If
'Exit Sub
'ErrTrap:
'    If BeginTrans = True Then
'        BeginTrans = False
'        Cn.RollbackTrans
'    End If
'    If Err.Number = -2147217900 Then
'        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
'        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
'        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Exit Sub
'    End If
'    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'End Sub
'Private Sub Del_TransAction()
'Dim RsTemp As ADODB.Recordset
'Dim Msg As String
'Dim StrSQL As String
'On Error GoTo ErrTrap
'If XPTxtMaintanenceID.text <> "" Then
'Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
'Msg = Msg + (XPTxtMaintanenceID.text) & Chr(13)
'Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
'If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
'    '«· √ﬂœ √‰Â ·„ Ì „ «” »œ«· √Ì ﬁÿ⁄Â ›Ì Â–Â «·⁄„·Ì…
'    StrSQL = "select * From  Transactions where MaintenanceID=" & Val(Rs("MaintananceID").Value)
'    Set RsTemp = New ADODB.Recordset
'    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RsTemp.EOF Or RsTemp.BOF) Then
'        Msg = "·ﬁœ  „ «” »œ«· √Õœ «·ﬁÿ⁄ ›Ì Â–Â «·⁄„·Ì… " & Chr(13)
'        Msg = Msg + "ÊÕ–› Â–Â «·⁄„·Ì… ”ÌƒœÌ ≈·Ï Õ–› »Ì«‰«  ⁄„·Ì… «·«” »œ«·" & Chr(13)
'        Msg = Msg + "Â·  —€» ›Ì Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…"
'        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
'             If Not Rs.RecordCount < 1 Then
'                Rs.Delete
'                StrSQL = "delete From Transactions where MaintenanceID=" & Val(XPTxtMaintanenceID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
'                Rs.MoveFirst
'                If Rs.RecordCount < 1 Then
'                    clear_all Me
'                    TxtModFlg_Change
'                    XPTxtCurrent.Caption = 0
'                    XPTxtCount.Caption = 0
'                Else
'                    Retrive
'                End If
'            End If
'        End If
'    Else
'        If Not Rs.RecordCount < 1 Then
'            Rs.Delete
'            Rs.MoveFirst
'            If Rs.RecordCount < 1 Then
'                clear_all Me
'                TxtModFlg_Change
'                XPTxtCurrent.Caption = 0
'                XPTxtCount.Caption = 0
'            Else
'                Retrive
'            End If
'        End If
'    End If
'End If
'Else
'    clear_all Me
'    Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
'    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'    TxtModFlg_Change
'    Exit Sub
'End If
'TxtModFlg_Change
'Exit Sub
'ErrTrap:
'If Err.Number = -2147217887 Then
'    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
'    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
'            vbExclamation, App.Title
'    Rs.CancelUpdate
'End If
'End Sub
'Private Sub Undo()
'On Error GoTo ErrTrap
'Select Case TxtModFlg.text
'    Case "N"
'         clear_all Me
'         Me.TxtModFlg.text = "R"
'         XPBtnMove_Click (1)
'    Case "E"
'         Rs.Find "MaintananceID='" & Val(XPTxtMaintanenceID.text) & "'", , adSearchForward, adBookmarkFirst
'         If Rs.EOF Or Rs.BOF Then
'            Me.TxtModFlg.text = "R"
'            Exit Sub
'        End If
'         Retrive
'         Me.TxtModFlg.text = "R"
'End Select
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub AddTip()
'Dim Wrap As String
'On Error GoTo ErrTrap
'Set TTP = New clstooltip
'Wrap = Chr(13) + Chr(10)
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(0), _
'    "ÃœÌœ ..." & Wrap & _
'    "·«÷«›… »Ì«‰«  ⁄„·Ì… ’Ì«‰… ÃœÌœ…" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(7), _
'    "ÿ»«⁄… ..." & Wrap & _
'    "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & _
'    " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(1), _
'    " ⁄œÌ· ..." & Wrap & _
'    "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(2), _
'    "Õ›Ÿ ..." & Wrap & _
'    "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·’Ì«‰…" & Wrap & _
'     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(3), _
'    " —«Ã⁄ ..." & Wrap & _
'    "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
'     "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(4), _
'    "Õ–› ..." & Wrap & _
'    "·Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(5), _
'    "»ÕÀ ..." & Wrap & _
'    "···»ÕÀ ⁄‰ ⁄„·Ì… ’Ì«‰…" & Wrap & _
'    "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl Cmd(6), _
'    "Œ—ÊÃ ..." & Wrap & _
'    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl XPBtnMove(1), _
'    "«·√Ê· ..." & Wrap & _
'    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
''With TTP
''   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
''   .MaxWidth = 4000
''   .VisibleTime = 9000
''   .DelayTime = 600
''   .AddControl CmdReplace, _
''    "«” »œ«· ..." & Wrap & _
''    "·«” »œ«· ﬁÿ⁄…  »⁄ «·÷„«‰" & Wrap & _
''    " ›ﬁÿ ≈÷€ÿ Â‰«", True
''End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl XPBtnMove(0), _
'    "«·”«»ﬁ ..." & Wrap & _
'    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl XPBtnMove(3), _
'    "«· «·Ì ..." & Wrap & _
'    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl XPBtnMove(2), _
'    "«·√ŒÌ— ..." & Wrap & _
'    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
'    " ›ﬁÿ ≈÷€ÿ Â‰«", True
'End With
'With TTP
'   .Create Me.hwnd, "⁄„·Ì«  «·’Ì«‰…", 1, 15204351, -2147483630
'   .MaxWidth = 4000
'   .VisibleTime = 9000
'   .DelayTime = 600
'   .AddControl CmdHelp, _
'    "„”«⁄œ… ..." & Wrap & _
'    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
'    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
'    "≈÷€ÿ Â‰«" & Wrap, True
'End With
'
'Exit Sub
'ErrTrap:
'End Sub
'Public Sub Retrive(Optional LngID As Long = 0)
'
'Dim RsDetails As ADODB.Recordset
'Dim RsReplace As ADODB.Recordset
'Dim StrSQL As String
'
'On Error GoTo ErrTrap
'
'If Rs.RecordCount < 1 Then
'    XPTxtCurrent.Caption = 0
'    XPTxtCount.Caption = 0
'    Exit Sub
'End If
'If Rs.EOF Or Rs.BOF Then
'    Exit Sub
'End If
'If LngID <> 0 Then
'    Rs.Find "MaintananceID=" & LngID, , adSearchForward, adBookmarkFirst
'    If Rs.EOF Or Rs.BOF Then
'        Exit Sub
'    End If
'End If
'XPTxtMaintanenceID.text = IIf(IsNull(Rs("MaintananceID").Value), "", (Rs("MaintananceID").Value))
'DBCboClientName.BoundText = IIf(IsNull(Rs("CusID").Value), "", Rs("CusID").Value)
'Me.DCboUserName.BoundText = IIf(IsNull(Rs("UserID").Value), "", Rs("UserID").Value)
'XPDtbGoInDtae.Value = IIf(IsNull(Rs("DateGoIN").Value), Date, Rs("DateGoIN").Value)
'XPDtbGoOutDtae.Value = IIf(IsNull(Rs("DateGoOUT").Value), Date, Rs("DateGoOUT").Value)
'CboMaintenanceType.ListIndex = IIf(IsNull(Rs("MType").Value), 0, Rs("MType").Value)
'Me.DcboEmp.BoundText = IIf(IsNull(Rs("EmpID").Value), "", Rs("EmpID").Value)
'Me.DCboStoreName.BoundText = IIf(IsNull(Rs("StoreID").Value), "", Rs("StoreID").Value)
'
'Fg.Rows = 2
'Fg.Clear flexClearScrollable, flexClearEverything
'StrSQL = "SELECT TblItems.HaveSerial,* FROM TblItems INNER JOIN TblMainteneceDetails " & _
'"ON TblItems.ItemID = TblMainteneceDetails.ItemID"
'StrSQL = "SELECT dbo.TblItems.HaveSerial, dbo.TblMainteneceDetails.*, " & _
'"dbo.TblCustemers.CusName "
'StrSQL = StrSQL + " FROM  dbo.TblItems INNER JOIN dbo.TblMainteneceDetails " & _
'"ON dbo.TblItems.ItemID = dbo.TblMainteneceDetails.ItemID LEFT OUTER JOIN "
'StrSQL = StrSQL + " dbo.TblCustemers ON dbo.TblMainteneceDetails.SupID = dbo.TblCustemers.CusID "
'StrSQL = StrSQL + " Where MaintananceID=" & Val(Rs("MaintananceID").Value)
'
'Set RsDetails = New ADODB.Recordset
'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'If Not (RsDetails.EOF Or RsDetails.BOF) Then
'    Fg.Rows = RsDetails.RecordCount + 1
'    For Num = 0 To RsDetails.RecordCount - 1
'        Fg.Cell(flexcpPicture, Num + 1, Fg.ColIndex("Replace")) = ""
'        Fg.Cell(flexcpData, Num + 1, Fg.ColIndex("Replace")) = ""
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").Value))
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").Value))
'        If (RsDetails("HaveSerial").Value) = True Then
'            Fg.Cell(flexcpChecked, Num + 1, Fg.ColIndex("HaveSerial")) = flexChecked
'        Else
'            Fg.Cell(flexcpChecked, Num + 1, Fg.ColIndex("HaveSerial")) = flexUnchecked
'        End If
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").Value))
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("EmpDes")) = IIf(IsNull(RsDetails("SupDeci")), "", Trim(RsDetails("SupDeci").Value))
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", Trim(RsDetails("Quantity").Value))
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("TicketNO")) = IIf(IsNull(RsDetails("TicketNO")), "", Trim(RsDetails("TicketNO").Value))
'        Fg.TextMatrix(Num + 1, Fg.ColIndex("EmpNotes")) = IIf(IsNull(RsDetails("EmpNotes")), "", Trim(RsDetails("EmpNotes").Value))
'
'        If IsNull(RsDetails("SupID").Value) Then
'            Fg.TextMatrix(Num + 1, Fg.ColIndex("SupManName")) = ""
'            Fg.Cell(flexcpData, Num + 1, Fg.ColIndex("SupManName")) = ""
'        Else
'            Fg.TextMatrix(Num + 1, Fg.ColIndex("SupManName")) = RsDetails("CusName").Value
'            Fg.Cell(flexcpData, Num + 1, Fg.ColIndex("SupManName")) = RsDetails("SupID").Value
'        End If
'        If Not IsNull(RsDetails("RetrunDate").Value) Then
'            Fg.TextMatrix(Num + 1, Fg.ColIndex("GoOutDate")) = RsDetails("RetrunDate").Value
'        End If
'
'        RsDetails.MoveNext
'    Next Num
'    Fg.AutoSize 0, Fg.Cols - 1, False
'End If
'XPTxtCurrent.Caption = Rs.AbsolutePosition
'XPTxtCount.Caption = Rs.RecordCount
'Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub XPTab301_Click()
'On Error GoTo ErrTrap
'If Me.TxtModFlg.text <> "R" Then
'    If XPTab301.CurrTab = 0 Then
'        XPBtnAdd.Enabled = True
'        XPBtnRemove.Enabled = True
'    Else
'        XPBtnAdd.Enabled = False
'        XPBtnRemove.Enabled = False
'    End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub PrintingData()
'On Error GoTo ErrTrap
'Dim ShowType As Boolean
'ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)
'If ShowType = True Then
'    If XPTxtMaintanenceID.text <> "" Then
'        Set MaintenReport = New ClsMaintananceReport
'        MaintenReport.MaintenanceDataShort XPTxtMaintanenceID.text
'    End If
'Else
'    If XPTxtMaintanenceID.text <> "" Then
'        Set MaintenReport = New ClsMaintananceReport
'        MaintenReport.MaintenanceData XPTxtMaintanenceID.text
'    End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
''Dim IntResult As String
''On Error GoTo ErrTrap
''If Me.TxtModFlg.text <> "R" Then
''Select Case Me.TxtModFlg.text
''    Case "N"
''        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
''        StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
''        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
''        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
''        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
''        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
''    Case "E"
''        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
''        StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
''        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
''        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
''        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
''        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
''End Select
''IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
''Select Case IntResult
''    Case vbYes
''        Cancel = True
''        SaveData
''    Case vbCancel
''        Cancel = True
''End Select
''End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ErrTrap
'If KeyCode = vbKeyReturn Then
'    If Me.TxtModFlg.text = "R" Then
'        Cmd_Click (0)
'    Else
'        SendKeys "{TAB}"
'    End If
'End If
'If KeyCode = vbKeyF12 Then
'    If Cmd(0).Enabled = False Then Exit Sub
'    Cmd_Click (0)
'End If
'If KeyCode = vbKeyF11 Then
'    If Cmd(1).Enabled = False Then Exit Sub
'    Cmd_Click (1)
'End If
'If KeyCode = vbKeyF10 Then
'    If Cmd(2).Enabled = False Then Exit Sub
'    Cmd_Click (2)
'End If
'If KeyCode = vbKeyF9 Then
'    If Cmd(3).Enabled = False Then Exit Sub
'    Cmd_Click (3)
'End If
'If KeyCode = vbKeyF8 Then
'    If Cmd(4).Enabled = False Then Exit Sub
'    Cmd_Click (4)
'End If
'If KeyCode = vbKeyF3 Then
'    If Cmd(5).Enabled = False Then Exit Sub
'    Cmd_Click (5)
'End If
'If KeyCode = vbKeyF6 Then
'    If Cmd(7).Enabled = False Then Exit Sub
'    Cmd_Click (7)
'End If
'If KeyCode = vbKeyF2 Then
'    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
'        XPBtnAdd_Click
'    End If
'End If
'If KeyCode = vbKeyF3 Then
'    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
'        XPBtnRemove_Click
'    End If
'End If
'If KeyCode = vbKeyF5 Then
'    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
'
'    End If
'End If
''If Shift = 2 Then
''    XPTab301.SetFocus
''    If KeyCode = vbKeyTab Then
''        If XPTab301.CurrTab = 0 Then
''            XPTab301.CurrTab = 1
''            If XPChkPayType(0).Enabled = True Then
''                XPChkPayType(0).SetFocus
''            End If
''        Else
''            XPTab301.CurrTab = 0
''            FG.SetFocus
''        End If
''    End If
''End If
'
'Exit Sub
'ErrTrap:
'End Sub
'
'
'
'
'Private Sub DBCboClientName_Change()
'On Error GoTo ErrTrap
'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'If DBCboClientName.BoundText <> "" Then
'    If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
'        CboPayMentType.Locked = True
'        CboPayMentType.ListIndex = 0
'    Else
'        CboPayMentType.Locked = False
'    End If
'End If
'End If
'Exit Sub
'ErrTrap:
'End Sub
'Private Sub DBCboClientName_Click(Area As Integer)
'DBCboClientName_Change
'End Sub
'Private Sub FillItemData()
'On Error GoTo ErrTrap
'Dim Dcombos As New ClsDataCombos
'Dcombos.GetItemsCodes Me.DCboItemsCode
'Dcombos.GetItemsNames Me.DCboItemsName
'Set cSearchDcbo(2) = New clsDCboSearch
'Set cSearchDcbo(2).Client = Me.DCboItemsCode
'Set cSearchDcbo(3) = New clsDCboSearch
'Set cSearchDcbo(3).Client = Me.DCboItemsName
'''Õ«·… «·’‰›
''With CboItemCase
''    .AddItem "ÃœÌœ"
''    .AddItem "„” ⁄„·"
''End With
'Exit Sub
'ErrTrap:
'End Sub
'
'
'Private Sub LoadTBR()
'With Me.TBar
'    .Buttons.Clear
'    .AllowCustomize = False
'    .Appearance = ccFlat
'    .BorderStyle = ccNone
'    .Style = tbrFlat
'    .TextAlignment = tbrTextAlignBottom
'    Set .ImageList = MDIFrmMain.ImgLstTree
'    .Buttons.Add , "RemoveRow", , , "Minus"
'
'End With
'End Sub
'
'Private Function CheckItemInv(LngItemID As Long, StrItemSerial As String, LngTransID As Long) As Boolean
'Dim StrSQL As String
'Dim Rs As ADODB.Recordset
'
'StrSQL = "select * From QryGuarantee where Item_ID=" & LngItemID
'StrSQL = StrSQL + " and ItemSerial='" & StrItemSerial & "'"
'StrSQL = StrSQL + " AND Transaction_ID='" & LngTransID & "'"
'StrSQL = StrSQL + " AND Transaction_Type=2"
'StrSQL = StrSQL + " AND CusID=" & DBCboClientName.BoundText
'Set Rs = New ADODB.Recordset
'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'If Rs.EOF Or Rs.BOF Then
'    CheckItemInv = False
'Else
'    CheckItemInv = True
'End If
'Rs.Close
'Set Rs = Nothing
'End Function
'
'
'
'
'
Private Sub Form_Load()

End Sub
