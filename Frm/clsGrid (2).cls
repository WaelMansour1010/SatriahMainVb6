VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Updated Function

class Name ClsGrid

 Public Function CheckDataEntered() As Boolean
    Dim Msg As String
    Dim StrCurrentItemName As String
    Dim StrCurrentItemCode As String
    Dim LngRowNum As Long, LngSearchResault As Long
    Dim StrTemp As String
    Dim LngItemID As Long
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim IntRes As Integer
    Dim StrItemSerial As String
    Dim RsTest As ADODB.Recordset
    Dim DblOutQty As Double
    Dim LngColorID As Long
    Dim LngClassId As Long
    Dim LngUnitID As Long

    Dim StrItemSize As String
'fillgrid (val(Me.STORENAME))
    '-------------------------------------------------
    'Start
    '-------------------------------------------------
    CheckDataEntered = False
'Debug.Print val(Me.storename.BoundText)
'GheckLinkItem
    '«·ﬂ‘› ⁄·Ï «‰ «·›« Ê—…  Õ ÊÏ ⁄·Ï √’‰«›
    If Me.Grid.FixedRows = Me.Grid.Rows Or ModFgLib.ItemsInGrid(Me.Grid, Me.Grid.ColIndex("Code")) = -1 Then
  If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ √’‰«› ›Ï  Â–« «·”‰œ ...!!!"
   Else
   Msg = "No Items In Voucher ...!!!"
   
   End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Screen.MousePointer = vbDefault
        Exit Function
    End If



'*****************************
'If SystemOptions.WorkWithLINKEDiTEMS = True Then
'    With Me.Grid

'        For LngRowNum = Me.Grid.FixedRows To Me.Grid.Rows - 1
'LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Name")))
' StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
 
' If GheckLinkItem(LngItemID, val(val(Me.storename.BoundText))) = False Then
'                 If SystemOptions.UserInterface = ArabicInterface Then
'                    Msg = " «·’‰›   " & StrCurrentItemName & Chr(13)
'                Else
'                     Msg = " Items   " & StrCurrentItemName & Chr(13)
'                End If
 
'                    Msg = Msg + StrCurrentItemName & Chr(13)
'                            If SystemOptions.UserInterface = ArabicInterface Then
'                              Msg = Msg + "€Ì— „—»Êÿ »Â–« «·„Œ“‰"
'                Else
'                     Msg = " Not Lnked With This Items   " & Chr(13)
'                End If
 
'          MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                    FocusGridOnCell LngRowNum, Grid.ColIndex("Count")
'                    Screen.MousePointer = vbDefault
'                    Exit Function
' End If
              '
                    
'    Next LngRowNum
    
'    End With
' End If
    
'*********************


    If Me.RestrictedAssbliedItemID <> 0 Then
        If CheckRestrictedItem = False Then
            CheckDataEntered = False
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If

    With Me.Grid

        For LngRowNum = Me.Grid.FixedRows To Me.Grid.Rows - 1

            If Val(.TextMatrix(LngRowNum, .ColIndex("Code"))) <> 0 Then
                StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
LngItemID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
  LngUnitID = Val(Grid.Cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")))
                If LngUnitID = 0 Then
                 GetDefaultItemUnit LngItemID, LngUnitID
                 Grid.TextMatrix(LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 Grid.Cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 
                End If
               
'---------------«· «ﬂœ „‰ «‰ „œ… «·⁄—÷ ﬁÌ„… —ﬁ„Ì…
If Grid.ColIndex("PODaysSub") <> -1 Then
             If Not IsNumeric(Val((Grid.TextMatrix(LngRowNum, Grid.ColIndex("PODaysSub"))))) Then
                                      Msg = " «—ÌŒ «· ”·Ì„  " & Chr(13)
                                      Msg = Msg + StrCurrentItemName & Chr(13)
                                      Msg = Msg + "€Ì— —ﬁ„Ì…"
                                      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                      FocusGridOnCell LngRowNum, Grid.ColIndex("PODaysSub")
                                      Screen.MousePointer = vbDefault
                                      Exit Function
            End If
End If
                '----------«· «ﬂœ „‰ ≈œŒ«· ﬂ„Ì… «·’‰›
                If Val(.TextMatrix(LngRowNum, .ColIndex("Count"))) <= 0 And Me.GridTrans <> NewGard Then
                    Msg = "„‰ ›÷·ﬂ ÌÃ»  ÕœÌœ «·ﬂ„Ì… «·Œ«’… »«·’‰›" & Chr(13)
                    Msg = Msg + StrCurrentItemName & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    FocusGridOnCell LngRowNum, Grid.ColIndex("Count")
                    Screen.MousePointer = vbDefault
                    Exit Function
                    '-------«· «ﬂœ „‰  ≈œŒ«· ”⁄— «·’‰›
                    '
                ElseIf Val(.TextMatrix(LngRowNum, .ColIndex("Price"))) <= 0 And (Me.GridTrans <> INVENTORYIN And Me.GridTrans <> ProductionOrder And Me.GridTrans <> ProductionMaterialReciveVoucher And Me.GridTrans <> InventoryOut And Me.GridTrans <> OpeningBalance And Me.GridTrans <> purchaseorderrequest And Me.GridTrans <> NewGard And Me.GridTrans <> internalorder And Me.GridTrans <> BookInventories And Me.GridTrans <> MoveItems And Me.GridTrans <> salespricelistRequest And Me.GridTrans <> ShipmentOrder And GridTrans <> ShipmentRegisteration And Me.GridTrans <> ShipmentRecieveVoucher And Me.GridTrans <> RecervieProductionVoucherNew And Me.GridTrans <> 24 And Me.GridTrans <> purchaserequest) Then
                    'Me.GridTrans <> PurchaseTransaction And
 'And Me.GridTrans <> Inventoryin
 
                    If Grid.ColIndex("DiscountType") <> -1 Then
                        
                        If Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) <> 4 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
                        Else
                        Msg = "Must Enter Price For Item" & Chr(13)
                        End If
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Price")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Price")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If

                    Else
           '             Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
                      If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
                        Else
                        Msg = "Must Enter Price For Item" & Chr(13)
                        End If
                        
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        Grid.Row = LngRowNum
                        Grid.Col = Grid.ColIndex("Price")
                        Grid.ShowCell LngRowNum, Grid.ColIndex("Price")
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Grid.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If

                    '-------«· «ﬂœ „‰ ≈œŒ«· ‰Ê⁄ «·Œ’„ ⁄·Ï «·»—‰«„Ã
                ElseIf Grid.ColIndex("DiscountType") <> -1 Then

                    If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                          Msg = "„‰ ›÷·ﬂ Õœœ ‰Ê⁄ «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                        Else
                        Msg = "Must Enter Discount Type For Item" & Chr(13)
                        End If
                        
                        
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Grid.Row = LngRowNum
                        Grid.Col = Grid.ColIndex("DiscountType")
                        Grid.ShowCell LngRowNum, Grid.ColIndex("DiscountType")
                        Grid.SetFocus
                        Exit Function
                        '-------«· «ﬂœ „‰ ≈œŒ«· ﬁÌ„… «·Œ’„ „Ÿ»Êÿ
                    ElseIf (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 2 Or (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 3 Then

                        If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal"))) = "" Then
                       
                                             If SystemOptions.UserInterface = ArabicInterface Then
                           Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                        Else
                        Msg = "Must Enter Discount Value For Item" & Chr(13)
                        End If
                        
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("DiscountVal")
                            Screen.MousePointer = vbDefault
                            Exit Function
                        ElseIf Not IsNumeric((Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal")))) Then
                        
                       '     Msg = "ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                                             If SystemOptions.UserInterface = ArabicInterface Then
                           Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                        Else
                        Msg = "Must Enter Discount Value For Item" & Chr(13)
                        End If
                       
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = Msg + "€Ì— —ﬁ„Ì…"
                        Else
                        Msg = "Not Numeric" & Chr(13)
                        End If
                       
                       
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("DiscountVal")
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                    End If
                End If

                If Me.GridTrans <> ShowPrice Then

                    '-------„—«Ã⁄… »Ì«‰«  «·”Ì—Ì«· «·„œŒ·
                    If .Cell(flexcpChecked, LngRowNum, .ColIndex("HaveSerial")) = flexChecked Then

                        '›Ï Õ«·… «‰ ÌﬂÊ‰ «·’‰› »‰Ÿ«„ «·”Ì—Ì«· ·«»œ „‰ «· «ﬂœ „‰ «·”Ì—Ì«· ﬁœ  „ ≈œŒ«·Â
                        If Trim$(.TextMatrix(LngRowNum, .ColIndex("Serial"))) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·’‰› " & Chr(13)
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Msg = Msg + "·Â ”Ì—Ì«·" & Chr(13)
                            Msg = Msg + "„‰ ›÷·ﬂ √œŒ· «·”Ì—Ì«· «·Œ«’ »Â"
                        Else
                        
                        Msg = "Item " & Chr(13)
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Msg = Msg + "Have Serial" & Chr(13)
                            Msg = Msg + "Please Enter Serial"
                        End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("Serial")
                            Screen.MousePointer = vbDefault
                            Exit Function
                            '›Ï ‰Ÿ«„ «·”Ì—Ì«· ·«»œ «‰ ÌﬂÊ‰ ﬂ· —ﬁ„ ›Ï Œ«‰… „‰›’·…
                        ElseIf Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Count"))) > 1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·’‰› " & Chr(13)
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Msg = Msg + "·Â ”Ì—Ì«·" & Chr(13)
                            Msg = Msg + "·« Ì„ﬂ‰  ”ÃÌ· √ﬂÀ— „‰ ﬁÿ⁄… Ê«Õœ…" & Chr(13)
                            Msg = Msg + "Ì„ﬂ‰ﬂ  ”ÃÌ· ﬂ· ﬁÿ⁄… ›Ì ’› „” ﬁ·"
                        Else
                        Msg = "Item " & Chr(13)
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Msg = Msg + "have Serial" & Chr(13)
                            Msg = Msg + "more item " & Chr(13)
                            Msg = Msg + "Ì„ﬂ‰ﬂ  ”ÃÌ· ﬂ· ﬁÿ⁄… ›Ì ’› „” ﬁ·"
                        
                        End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Count")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Count")
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                            'ÿÊ· «·—ﬁ„ «·”Ì—Ì«· ·«Ì“Ìœ ⁄‰ 50 Õ—›
                        ElseIf Len(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) > 50 Then
                            Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                            Msg = Msg & StrCurrentItemName & Chr(13)
                            Msg = Msg + "€Ì— „ﬁ»Ê·" & Chr(13)
                            Msg = Msg + "ÌÃ» √·« Ì“Ìœ ÿÊ· «·”Ì—Ì«· ⁄‰ 50 Õ—› √Ê —ﬁ„"
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Serial")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Serial")
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                        ElseIf .ColIndex("guaranteeTime") <> -1 Then

                            If .TextMatrix(LngRowNum, .ColIndex("guaranteeTime")) = "" Then
                            .TextMatrix(LngRowNum, .ColIndex("guaranteeTime")) = 1
                                If Not (IsNumeric(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) < 1 Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) > 36 Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) <> Int(Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Then
                                    Msg = "„œ… «·÷„«‰ «·Œ«’… »«·’‰›" & Chr(13)
                                    Msg = Msg + StrCurrentItemName & Chr(13)
                                    Msg = Msg + "€Ì— ’«·Õ…" & Chr(13)
                                    Msg = Msg + "√œŒ· „œ…   —«ÊÕ »Ì‰ 1 : 36 ‘Â—" & Chr(13)
                                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    Grid.Row = LngRowNum
                                    Grid.Col = Grid.ColIndex("guaranteeTime")
                                    Grid.ShowCell LngRowNum, Grid.ColIndex("guaranteeTime")
                                    Grid.SetFocus
                                    Screen.MousePointer = vbDefault
                                    Exit Function
                                End If
                            End If
                        End If
                    End If

                    '--------------------------------------
                End If
            End If

        Next LngRowNum

        '--------------------------
        If Me.WorkOrderID <> 0 Then
            If Me.RestrictedAssbliedItemID <> 0 Then
                If CheckRestrictedItemQty(Me.WorkOrderID, Me.RestrictedAssbliedItemID) = False Then
                    CheckDataEntered = False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        End If

        '--------------------------
        For LngRowNum = 1 To Grid.Rows - 1

            If Trim$(Me.Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) <> "" Then
                LngSearchResault = Grid.FindRow(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")), , 4)

                If LngSearchResault <> LngRowNum And LngSearchResault > 0 Then
                    StrTemp = Grid.Cell(flexcpTextDisplay, LngSearchResault, Grid.ColIndex("Name"))
                    Msg = " „ «œŒ«· «·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                    Msg = Msg + StrCurrentItemName & Chr(13)
                    Msg = Msg + "›Ï «·”ÿ—:" & LngRowNum
                    Msg = Msg + "ﬁ»· –·ﬂ «·Ì «·’‰›" & Chr(13)
                    Msg = Msg + StrTemp & Chr(13)
                    Msg = Msg + "›Ï «·”ÿ—:" & LngSearchResault
                    Msg = Msg + " √ﬂœ „‰ «·”Ì—Ì«· «·Œ«’ »«·√’‰«›"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    FocusGridOnCell LngRowNum, .ColIndex("Serial")
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If

        Next LngRowNum

        '---------------------------------------------------------------------------
        '›Ï Õ«·… «‰  ﬂÊ‰ ›« Ê—… „‘ —Ì«  «Ê „— Ã⁄ „»Ì⁄« 
        '›«‰ «·»—‰«„Ã Ì„‰⁄  ﬂ—«— «·”Ì—Ì«· „⁄ «·„ÊÃÊœ ›Ï
        'ﬁ«⁄œ… «·»Ì«‰« 
        If (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = OpeningBalance) Then

            For LngRowNum = 1 To Grid.Rows - 1

                If Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")) <> "" Then
                    StrItemSerial = Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")))
                    StrSQL = "Select * From Transaction_Details where ItemSerial='" & StrItemSerial & "'"
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If RsTemp.RecordCount > 0 Then
                        LngItemID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
                        Set RsRepeat = New ADODB.Recordset

                        If Me.TxtModFlag.Text = "N" Then
                            Set RsRepeat = GetItemQuantityStock(LngItemID, Val(Me.StoreName.BoundText), , , , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                        ElseIf Me.TxtModFlag.Text = "E" Then
                            Set RsRepeat = GetItemQuantityStock(LngItemID, Val(Me.StoreName.BoundText), , Val(Me.TxtInvID.Text), , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                        End If

                        If Not (RsRepeat.EOF Or RsRepeat.BOF) Then
                        'WaelEdit
                        'OldCode
                             'If RsRepeat("Qty").value > 0 Then
                         'NewCode
                            If RsRepeat("TotalQty").Value > 0 Then
                            'NewCode
                                Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                                Msg = Msg + StrCurrentItemName & Chr(13)
                                Msg = Msg + "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                                Msg = Msg + "„ÊÃÊœ ‰›” «·”Ì—Ì«· ·œÌ‰« ›Ï «·„Œ“‰ "

                                If Not StoreName Is Nothing Then
                                    Msg = Msg & "[" & StoreName.Text & "]"
                                End If

                                Msg = Msg & Chr(13) & "------------------------------------"
                                Msg = Msg & Chr(13) & "Õ«Ê· «‰  ” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ ”Ì—Ì«· ·„⁄—›… «·„“Ìœ"
                                Msg = Msg & Chr(13) & "Â·  —Ìœ «‰ ÌﬁÊ„ «·»—‰«„Ã »«·»ÕÀ ·ﬂ ⁄‰  Â–« «·”Ì—Ì«· «·¬‰...!!!"
                                IntRes = MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

                                If IntRes = vbYes Then
                                    'Open the Search for serial screen
                                    OpenScreen CheckItemSerial, LngItemID, StrItemSerial
                                ElseIf IntRes = vbNo Then
                                    FocusGridOnCell LngRowNum, .ColIndex("Serial")
                                End If

                                Screen.MousePointer = vbDefault
                                Exit Function
                            End If
                        End If

                        RsRepeat.Close
                    End If

                    RsTemp.Close
                End If

            Next LngRowNum

        End If

        '--------------------------------------------------------------------------
        '--------------------------------------------------------------------------
        '›Ï Õ«·… «‰  ﬂÊ‰ ›« Ê—… „»Ì⁄«  «Ê „— Ã⁄ „‘ —Ì« 
        '·«»œ „‰ «· «ﬂœ „‰ «‰ —’Ìœ «·’‰› ›Ï «·„Œ“‰ «·„Õœœ Ì”„Õ
        'Ê«‰ «·”Ì—Ì«· «·„—«œ Œ—ÊÃÂ „ÊÃÊœ ›Ï «·„Œ“‰
        'ŒÿÊ«  «· «ﬂœ „‰ «·ﬂ„Ì…   „ ⁄·Ï „— Ì‰
        '---------
        '(«·ŒÿÊ… «·√Ê·Ï) («· «ﬂœ „‰ «‰ «·ﬂ„Ì… √”«”«  ”„Õ)
        '---------
        '(«·ŒÿÊ… «·À«‰Ì…) (»⁄œ «· «ﬂœ „‰ «‰  √”«”«  ”„Õ)
        '»«·‰”»… ··√’‰«› «· Ï  ⁄„· »«·”Ì—Ì«·
        'ÌÃ» «· «ﬂœ «‰ «·”Ì—Ì«· „ÊÃÊœ ›Ï «·„Œ“‰
        '-------------------------------------------------------------------------------------------------

        If SelectedIssueVoucher = True Then
            CheckDataEntered = True
            Exit Function
        End If

        If SystemOptions.TypicalProduction = False Then
            'If (Me.GridTrans = 2 Or Me.GridTrans = 14 Or Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = Destruction) Then
If (Me.GridTrans = 2 Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction) Then
                For LngRowNum = 1 To Grid.Rows - 1
                    CalculateInvProfit LngRowNum
                Next

            End If
 
        End If

        If (Me.GridTrans = 31 Or Me.GridTrans = 2 Or Me.GridTrans = MoveItems Or Me.GridTrans = 14 Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction) Then
            If (Me.GridTrans = 31 Or Me.GridTrans = 2 Or Me.GridTrans = 14 Or Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = Destruction) Then
                If SystemOptions.SysAllowStockNegative = False Then

                    For LngRowNum = 1 To Grid.Rows - 1
                
                        If Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType"))) = 0 Then '·«»œ Ê«‰ ÌﬂÊ‰ «·’‰› ”·⁄…
                            If Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")) <> "" Then
                                '-------------------------------------------------------------------
                                '«·»Ì«‰«  «·√”«”Ì… —ﬁ„ «·’‰› Ê «”„ «·’‰› Ê «·”Ì—Ì«·
                                LngItemID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
                                StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
                                StrCurrentItemCode = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Code"))
                                LngColorID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ColorID")))
                                StrItemSize = Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemSize"))
                                LngClassId = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ClassId")))

                                If Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) <> "" Then
                                    StrItemSerial = Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")))
                                Else
                                    StrItemSerial = ""
                                End If

                                DblOutQty = CheckTotalQty(LngRowNum, LngColorID, StrItemSize, LngClassId, StrItemSerial)
                       
                                '-------------------------------------------------------------------
                                If LngColorID = 0 Then LngColorID = 1
                                 If LngClassId = 0 Then LngClassId = 1
                                 If StrItemSize = "" Then StrItemSize = "1"
                                If Me.TxtModFlag.Text = "N" Then
                                    Set RsTest = GetItemQuantityStock(LngItemID, Val(Val(Me.StoreName.BoundText)), Me.DtpBillDate.Value, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                ElseIf Me.TxtModFlag.Text = "E" Then
                                    '·«ÕŸ «‰Â ›Ï Õ«·… «· ⁄œÌ· ·«»œ „‰ ≈—”«· —ﬁ„ «·›« Ê—… ··œ«·…
                        
                                    If (Me.GridTrans = InvoiceTransaction) Or (Me.GridTrans = Returntransaction) Then
                      
                                        Set RsTest = GetItemQuantityStock(LngItemID, Val(Me.StoreName.BoundText), Me.DtpBillDate.Value, Val(Me.TxtNots.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                        
                                        'Set RsTest = GetItemQuantityStock(LngItemID, val(Me.storename.BoundText), _
                                        'Me.DtpBillDate.value,  Val(Me.TxtInvID.text), , , StrItemSerial, True, LngColorID, StrItemSize)
                                    Else
                        
                                        Set RsTest = GetItemQuantityStock(LngItemID, Val(Me.StoreName.BoundText), Me.DtpBillDate.Value, Val(Me.TxtInvID.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                        
                                    End If
                        
                                End If

                                'Stop here
                                Screen.MousePointer = vbDefault
                    
                                If RsTest.EOF Or RsTest.BOF Then
                                    Msg = "«·’‰› : " & StrCurrentItemName & Chr(13)
                                    Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)

                                    If StrItemSerial <> "" Then
                                        Msg = Msg & "«·”Ì—Ì«· : " & StrItemSerial & Chr(13)
                                    End If
                        
                                    Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                                    Msg = Msg & "·« ÊÃœ „‰Â «Ï ﬂ„Ì… ›Ï «·„Œ“‰ : " & Me.StoreName.Text & Chr(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                                    Msg = Msg & "--------------------------" & Chr(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    FocusGridOnCell LngRowNum, .ColIndex("Name")
                                    TerminateRecordset RsTest
                                    Exit Function
                                ElseIf (Round(RsTest("totalqty").Value, 2) < Round(DblOutQty, 2)) And StrItemSerial = "" Then
                                    Msg = "«·’‰› : " & StrCurrentItemName & Chr(13)
                                    Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                                    Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                                    Msg = Msg & "≈Ã„«·Ï «·ﬂ„Ì… «·„ÊÃÊœ… „‰Â ›Ï «·„Œ“‰ : " & Me.StoreName.Text & Chr(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                                    Msg = Msg & "ÂÏ : " & RsTest("totalqty").Value & Chr(13)
                                    Msg = Msg & "«·ﬂ„Ì… «·„—«œ ’—›Â« : " & DblOutQty & Chr(13)
                                    Msg = Msg & "·«Ì„ﬂ‰ «·”„«Õ »Œ—ÊÃ Â–Â «·ﬂ„Ì….  " & Chr(13)
                                    Msg = Msg & "--------------------------" & Chr(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    FocusGridOnCell LngRowNum, .ColIndex("Count")
                                    TerminateRecordset RsTest
                                    Exit Function
                                ElseIf StrItemSerial <> "" Then
                                    '  RsTest.MoveFirst
                                    '  RsTest.find "ItemSerial='" & StrItemSerial & "'", , adSearchForward, adBookmarkFirst
                                    '      If RsTest.BOF Or RsTest.EOF Then
                                    '      Msg = "«·”Ì—Ì«· : " & StrItemSerial & Chr(13)
                                    '      Msg = Msg & "«·’‰› : " & StrCurrentItemName & Chr(13)
                                    '      Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                                    '      Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                                    '      Msg = Msg & "Â–« «·”Ì—Ì«· €Ì— „ÊÃÊœ »«·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                                    '      Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & Chr(13)
                                    '      Msg = Msg & "--------------------------" & Chr(13) & ""
                                    '      Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                                    '      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    '      FocusGridOnCell LngRowNum, .ColIndex("Count")
                                    '      TerminateRecordset RsTest
                                    '                            Exit Function
                                    '      ElseIf RsTest("Qty").value <= 0 Then
                                    '          Msg = "«·”Ì—Ì«· : " & StrItemSerial & Chr(13)
                                    '          Msg = Msg & "«·’‰› : " & StrCurrentItemName & Chr(13)
                                    '          Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                                    '          Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                                    '          Msg = Msg & "Â–« «·”Ì—Ì«· €Ì— „ÊÃÊœ »«·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                                    '          Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & Chr(13)
                                    ''          Msg = Msg & "--------------------------" & Chr(13) & ""
                                    '          Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                                    '          MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    '          FocusGridOnCell LngRowNum, .ColIndex("Count")
                                    '          TerminateRecordset RsTest
                                    '          Exit Function
                        
                                    '      End If
                    
                                End If
                
                            End If
                        End If
                        
                        
                
                    Next LngRowNum

                End If
            End If
        End If

    End With

    '-------------------------------------------------
    'END
    '-------------------------------------------------
    Screen.MousePointer = vbDefault
    CheckDataEntered = True
    Exit Function
End Function




