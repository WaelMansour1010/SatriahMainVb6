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
Option Explicit
Private m_GridTrans2 As GridTransType
Public WithEvents TxtLotNo As TextBox
Attribute TxtLotNo.VB_VarHelpID = -1
Public WithEvents TxtCustCode As TextBox 'ﬂÊœ «·⁄„Ì·
Attribute TxtCustCode.VB_VarHelpID = -1
Public bassprofit As Boolean
Public WithEvents Grid As VSFlexGrid
Attribute Grid.VB_VarHelpID = -1
Public WithEvents VatGrid As VSFlexGrid
Attribute VatGrid.VB_VarHelpID = -1
Public ReturnTyp As Integer
Public WithEvents CBoBasedON As ComboBox
Attribute CBoBasedON.VB_VarHelpID = -1
Public WithEvents TxtModFlag As VB.TextBox
Attribute TxtModFlag.VB_VarHelpID = -1
Public WithEvents TxtInventoryAction  As VB.TextBox
Attribute TxtInventoryAction.VB_VarHelpID = -1
Public WithEvents GrdTBar As MSComctlLib.Toolbar
Attribute GrdTBar.VB_VarHelpID = -1
Public txtTotal As TextBox
Public TxtValueAdded As TextBox
Public WithEvents Customer As DataCombo
Attribute Customer.VB_VarHelpID = -1
Public WithEvents Branch As DataCombo
Attribute Branch.VB_VarHelpID = -1
Dim cCompanyInfo As New ClsCompanyInfo
Public WithEvents CboDiscount_Type As ComboBox
Attribute CboDiscount_Type.VB_VarHelpID = -1
Public SelectedLotNo As String
Public WithEvents TxtDiscount_Val As TextBox
Attribute TxtDiscount_Val.VB_VarHelpID = -1
Public lblTotalSalesPrice As Label
Public txt_Currency_rate As TextBox

Public TxtValueCash As TextBox

Public TxtValueDelay As TextBox

Public TxtValuechque As TextBox

Public WithEvents TxtTaxValue As TextBox
Attribute TxtTaxValue.VB_VarHelpID = -1

Public WithEvents TxtAddTax As TextBox
Attribute TxtAddTax.VB_VarHelpID = -1

Public WithEvents TxtStampTax As TextBox
Attribute TxtStampTax.VB_VarHelpID = -1

Public WithEvents TxtServiceTax As TextBox
Attribute TxtServiceTax.VB_VarHelpID = -1
Public TxtFillData As TextBox


Public WithEvents StoreName As DataCombo
Attribute StoreName.VB_VarHelpID = -1

Public LblInvProfit As Label

Public LblItemsCount As Label

Public WithEvents DtpBillDate As DTPicker
Attribute DtpBillDate.VB_VarHelpID = -1

Public TxtInvID As TextBox

Public TxtNots As TextBox

Private FrmComment As FrmFlexQtyTip

Public WithEvents m_FrmItemsNewPrices As Form
Attribute m_FrmItemsNewPrices.VB_VarHelpID = -1

Public WithEvents LblTotalAll As Label
Attribute LblTotalAll.VB_VarHelpID = -1
Public WithEvents TxtItemCodeB    As TextBox
Attribute TxtItemCodeB.VB_VarHelpID = -1
Public WithEvents TxtItemCodeB1    As TextBox
Attribute TxtItemCodeB1.VB_VarHelpID = -1

Public WithEvents LblDiscountsTotal As Label
Attribute LblDiscountsTotal.VB_VarHelpID = -1

Public WithEvents LblCommision As Label
Attribute LblCommision.VB_VarHelpID = -1

Public WithEvents LblTotalQty As Label
Attribute LblTotalQty.VB_VarHelpID = -1

Public LblTaxSalesValue As Label

Public LblTaxAddValue As Label

Public LblTaxStampValue As Label

Public LblTaxServiceValue As Label
Public LBLGross As Label

'«·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
Public WithEvents DCboItemCode As DataCombo  'ﬂÊœ «·’‰›
Attribute DCboItemCode.VB_VarHelpID = -1

Public WithEvents DCboItemName As DataCombo  '«”„ «·’‰›
Attribute DCboItemName.VB_VarHelpID = -1

Public WithEvents DcboEmp As DataCombo  '«”„ «·’‰›
Attribute DcboEmp.VB_VarHelpID = -1
Public WithEvents TxtCashCustomerName As TextBox          '«”„ „Œ ’—
Attribute TxtCashCustomerName.VB_VarHelpID = -1

Public WithEvents CboItemCase As ComboBox    'Õ«·… «·’‰›
Attribute CboItemCase.VB_VarHelpID = -1

Public WithEvents TxtShortName As TextBox          '«”„ „Œ ’—
Attribute TxtShortName.VB_VarHelpID = -1


Public WithEvents TxtSerial As TextBox       '«·”Ì—Ì«·
Attribute TxtSerial.VB_VarHelpID = -1

Public WithEvents TxtQuantity As TextBox     '«·ﬂ„Ì…
Attribute TxtQuantity.VB_VarHelpID = -1

Public WithEvents TxtPrice As TextBox         '«·”⁄—
Attribute TxtPrice.VB_VarHelpID = -1

Public WithEvents CmdAddData As ISButton   '“— ≈÷«›…
Attribute CmdAddData.VB_VarHelpID = -1

Public WithEvents CmdAddSerialLIst As ISButton
Attribute CmdAddSerialLIst.VB_VarHelpID = -1


'Public TxtGuarante As TextBox  '„œ… «·÷„«‰
'Public CboDiscountType As  ComboBox    '‰Ê⁄ «·Œ’„
'Public TxtDiscountVal As TextBox  'ﬁÌ„… «·Œ’„
Private m_GridTrans As GridTransType
Dim cSearchDcbo(3) As clsDCboSearch

Private Enum EditDir
    EditDirVer
    EditDirHor
    EditDirFree
End Enum

Private Enum ItemRowToolTip
    ItemQtyToolTip
    ItemPriceToolTip
End Enum


Public Enum ItemsType
    ItemsAllType
    ItemsGoodType
    ItemsServiceType
End Enum

Private m_ShowQtyToolTip As Boolean

Private m_QtyToolTipShown As Boolean

Private LngCurrentItemID As Long

Private MyAddItemForm As FrmAddItem
Dim m_EditDir As EditDir
Dim RsItems As ADODB.Recordset

Private m_RestrictedAssbliedItemID As Long

Private m_WorkOrderID As Long
              Dim LastPurchaseDate As String
                           Dim LastPurchasePrice As Double
                           Dim LastPurchaseqty As Double
                           Dim AverageIssue As Double
                           Dim AverageIssueyraly As Double
                           Dim RequestLimit As Double
                            Dim Fromdate As Date
                        Dim AverageMonthly As Double
                        Dim ItemBalance As Double
                         Dim FirstPeriodDateInthisYear  As Date

Private QtyFromBarCode As String
Private CodeFromBarCode As String

Public Property Get GridTrans2() As GridTransType
    GridTrans2 = m_GridTrans2
End Property
Private Sub Customer_GotFocus()
  'Wael Add
    Customer.backcolor = vbYellow

End Sub

Private Sub Customer_KeyDown(KeyCode As Integer, Shift As Integer)
'Wael Add
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Customer_LostFocus()
   'Wael Add
    Customer.backcolor = vbWhite
End Sub


Private Sub TxtCustCode_GotFocus()
'Wael add
TxtCustCode.backcolor = vbYellow
End Sub

Private Sub TxtCustCode_KeyDown(KeyCode As Integer, Shift As Integer)
'Wael add
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtCustCode_LostFocus()
'Wael add
 TxtCustCode.backcolor = vbWhite
End Sub


'Wael  Add
 Public Property Let GridTrans2(ByVal vNewValue As GridTransType)

    If (vNewValue <> ReturnDestruction _
And vNewValue <> purchaserequest And vNewValue <> InventoryOut And vNewValue <> InvoiceTransaction And vNewValue <> InvoiceTransactionCompose _
  And vNewValue <> PurchaseTransaction And vNewValue <> Returntransaction And vNewValue <> GridTransType.ShowPrice And vNewValue <> Template And vNewValue <> Destruction And vNewValue <> ReturnSalling And vNewValue <> MoveItems And vNewValue <> OpeningBalance And vNewValue <> StockSettlement And vNewValue <> ProductionOrder And vNewValue <> NewGard And vNewValue <> ProductionMaterialReciveVoucher And vNewValue <> purchaseorderrequest And vNewValue <> internalorder And vNewValue <> INVENTORYIN And vNewValue <> salespricelistRequest And vNewValue <> salespricelist And vNewValue <> purchaseorder And vNewValue <> ShipmentOrder And vNewValue <> ShipmentRegisteration And vNewValue <> ShipmentRecieveVoucher And vNewValue <> purchaseOrderApproved And vNewValue <> RecervieProductionVoucherNew And vNewValue <> BookInventories And vNewValue <> NewGard And vNewValue <> internalissuerequesT) Then
    Else
        m_GridTrans2 = vNewValue
    End If

End Property

Function CheckMaultyItem(ByRef LngRow As Long, LngItemID As Long, Optional DefalutUnitID As Long, Optional ByRef DblQty As Double) As Boolean
Dim i As Long
CheckMaultyItem = False
With Grid
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("Code"))) = LngItemID And i <> LngRow Then
.TextMatrix(i, .ColIndex("Count")) = val(.TextMatrix(i, .ColIndex("Count"))) + DblQty
.RemoveItem LngRow
DblQty = val(.TextMatrix(i, .ColIndex("Count")))
LngRow = i
CheckMaultyItem = True
Exit Function
End If
Next i
End With
End Function

Function GetStoreID2(Optional ItemNameID As Double, Optional MaxNo As String) As Double
If ItemNameID = 0 Or MaxNo = "" Then Exit Function
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     StoreID2"
sql = sql & " From dbo.TblDefComItem"
sql = sql & " WHERE     (ItemNameID = " & ItemNameID & ") AND (MaxNo = N'" & MaxNo & "')"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetStoreID2 = IIf(IsNull(rs2("StoreID2").value), 0, rs2("StoreID2").value)
Else
GetStoreID2 = 0
End If
End Function
                         
                         
                         Sub RemoveItems()
Dim Number As Integer
Number = 0
Do
Number = Number + 1
If Number < Grid.Rows Then
   If Grid.Rows > 2 And Grid.TextMatrix(Number, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Number
    
    End If
 End If
Loop While Number < Grid.Rows
   If Grid.Rows - 1 > 1 And Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Grid.Rows - 1
    End If
End Sub
                        
Sub RemoveItemsx04122017()
Dim Number As Integer
Number = 0
Do
Number = Number + 1
If Number < Grid.Rows Then
   If Grid.Rows > 1 And Grid.TextMatrix(Number, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Number
    
    End If
 End If
Loop While Number < Grid.Rows
   If Grid.Rows - 1 > 1 And Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Grid.Rows - 1
    End If
End Sub

Public Function FindItemInGrid(LngRow As Long, _
                               LngItemID As Long) As Long

    'LngRow:-      the Row which the user edit
    'LngItemID:-   the Item ID we will search For it

    Dim LngFound As Long

    With Me.Grid
        LngFound = .FindRow(LngItemID, LngRow, .ColIndex("Name"), False, True)

        If LngFound = -1 Then
            'NOT Exist In the Grid
            FindItemInGrid = -1
        Else
            FindItemInGrid = LngFound
        End If

    End With

End Function

Private Sub CboDiscount_Type_Click()
SentTypeVAT
End Sub

Private Sub CboItemCase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
   
        SendKeys "{TAB}"
    End If
End Sub
 
 
 Private Sub CboItemCase_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then
   
        SendKeys "{TAB}"
    End If

End Sub
Sub ColorExperDate()
RemoveItems
 Dim i As Integer
If SystemOptions.AllowExperDateFIFO = True Then
With Grid
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
If .TextMatrix(i, .ColIndex("ExpiryDate")) = "" Then
.Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
ElseIf DateDiff("d", .TextMatrix(i, .ColIndex("ExpiryDate")), Date) >= 0 Then
.Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
End If
End If
Next i
End With
End If
End Sub
Private Sub Class_Initialize()
    m_EditDir = EditDirVer
    'Set m_MenuItemCart = MDIFrmMain.MnuItemTools_ItemCart
    'Set m_MenuItemData = MDIFrmMain.MnuItemTools_ItemData
    'Set m_MenuItemQty = MDIFrmMain.MnuItemTools_ItemQty
    'Set m_MenuItemSerial = MDIFrmMain.MnuItemTools_ItemSerial
End Sub

Public Sub Class_Terminate()

    If Not RsItems Is Nothing Then
        If RsItems.State = adStateOpen Then
            RsItems.Close
        End If

        Set RsItems = Nothing
    End If

    CloseMyAddItemForm
    CloseNewItemsPrices
  SaveMySetting
End Sub


 
Private Sub CmdAddData_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyReturn Then
    
        DCboItemCode.SetFocus
 

 
    End If

End Sub



Public Sub CmdAddData_Click() 'salim2
    '“— «·≈÷«›… ·‰ﬁ· »Ì«‰«  «·√’‰«› ≈·Ï «·ÃœÊ·
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer
    Dim LngRow As Long
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim StrItemDataMsg As String

    
    If DCboItemCode.Text = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
                        Else
                        Msg = "Enter item Code"
                        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemCode.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DCboItemName.Text = "" Then
  '      Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        
                                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
                        Else
                        Msg = "Enter item name"
                        End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If
'Wael Update

    If Me.GridTrans2 = InvoiceTransactionCompose Then
    If Customer.Text = "" Then
  '      Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        
                                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «”„ «·⁄„Ì·"
                        Else
                        Msg = "Enter customer name"
                        End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Customer.SetFocus
        SendKeys "{F4}"
        Exit Sub
        End If
    End If

    If CboItemCase.ListIndex = -1 Then
       ' Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
                        Else
                        Msg = "Enter item Status"
                        End If


        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboItemCase.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Not TxtSerial Is Nothing Then
        If Trim(Me.TxtSerial.Text) <> "" Then

            With Me.Grid
                StrSerial = Trim(Me.TxtSerial.Text)
                LngRow = .FindRow(StrSerial, .FixedRows, .ColIndex("Serial"), False, True)

                If LngRow <> -1 Then
                    Msg = "Â–« «·”Ì—Ì«· „œŒ· „”»ﬁ« ›Ï «·”ÿ— —ﬁ„ " & LngRow
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

            End With

        End If
    End If

    If val(TxtQuantity.Text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰›"
                        Else
                        Msg = "Enter item Qty"
                        End If
                        
   '     Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtQuantity.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(TxtQuantity.Text) Then
      '  Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        
                    If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                        Else
                        Msg = "Qty Must be Numeric"
                        End If
                        
                        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtQuantity.SetFocus
        Exit Sub
    End If

    If TxtPrice.Text = "" Then
        'Msg = "ÌÃ»  ÕœÌœ «·”⁄—"
        
                     If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «·”⁄— "
                        Else
                        Msg = "Price Must be Entered"
                        End If
                        
                        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtPrice.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(TxtPrice.Text) Then
        Msg = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                     If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·”⁄— ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                        Else
                        Msg = "Price Must be Numeric"
                        End If
                        
                        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtPrice.SetFocus
        Exit Sub
    End If

    If Me.TxtQuantity.Tag = "T" Then
        If Not Me.TxtSerial Is Nothing Then
            If Trim(Me.TxtSerial.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… «·”Ì—Ì«· ‰„»—...!!"
                
                                     If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÌÃ» ﬂ «»… «·”Ì—Ì«· ‰„»—...!! "
                        Else
                        Msg = "Serial Number Must be Entered"
                        End If
                        
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtSerial.SetFocus
                Exit Sub
            End If
        End If
    End If

    '-------------------------------------------------------------------------------------------------
    LngItemID = val(Me.DCboItemName.BoundText)
    'Check if this is Assblied Item
    'Or it is a Services Item
    RsItems.MoveFirst
    RsItems.find "ItemID=" & LngItemID, , adSearchForward, adBookmarkFirst

    If RsItems.BOF Or RsItems.EOF Then
       ' Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ"
        
                                     If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ"
                        Else
                        Msg = "Invalid Item Code"
                        End If
                        
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    ElseIf (RsItems("AssbliedItem").value = True Or RsItems("AssbliedItem").value = -1 Or RsItems("AssbliedItem").value = 1) Then
        '’‰› „Ã„⁄
    
    ElseIf (RsItems("ItemType").value = 1) Then
        '’‰› Œœ„…
    
    Else
        StrItemDataMsg = "«·’‰› : " & RsItems("ItemName").value
        StrItemDataMsg = StrItemDataMsg & CHR(13) & "—ﬁ„:" & RsItems("ItemID").value
        StrItemDataMsg = StrItemDataMsg & CHR(13) & "ﬂÊœ:" & RsItems("ItemCode").value

        '’‰› ⁄«œÏ
        '-------------------------------------------------------------------------------
        'ÌÃ» «· Õﬁﬁ „‰ «‰ «·ﬂ„Ì…  ”„Õ (›Ï Õ«·… ›Ê« Ì— «·„»Ì⁄«  «Ê „— Ã⁄ «·„‘ —Ì« ) ﬁ»· ≈÷«›… «·’‰›
        'Ê«‰ ≈⁄œ«œ«  «·»—‰«„Ã „⁄œ… ⁄·Ï  ‰Ÿ«„ «·Ã—œ «·›⁄·Ï
        ' ›Ï Õ«·… «‰ ÌﬂÊ‰ „— Ã⁄ „»Ì⁄«  «Ê ›« Ê—… „‘ —Ì«  ›ÌÃ» «· «ﬂœ „‰ „‰⁄  ﬂ—«—
        '«·”Ì—Ì«· ‰„»—
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Then

            '›« Ê—… »Ì⁄ «Ê „— Ã⁄ „‘ —Ì« 
            If SystemOptions.SysAllowStockNegative = False Then 'Must Check the Stock
            '    Set RsTemp = New ADODB.Recordset
            '    Set RsTemp = GetItemQuantityStock(LngItemID, val(val(Me.storename.BoundText)), Me.DtpBillDate.value, , , , Trim(Me.TxtSerial.text))

             '   If RsTemp.BOF Or RsTemp.EOF Or IsNull(RsTemp("totalqty").value) Then
             '       Msg = StrItemDataMsg
             '       Msg = Msg & "«·ﬂ„Ì… ·«  ”„Õ .. √Ê «‰ «·”Ì—Ì«· «·„œŒ· €Ì— ’ÕÌÕ"
             '       MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
             '       Exit Sub
             '   ElseIf RsTemp("totalqty").value < val(Me.TxtQuantity.text) Then
             '       Msg = StrItemDataMsg
             '       Msg = Msg & Chr(13) & "«·ﬂ„Ì… ·«  ”„Õ .."
             '       Msg = Msg & Chr(13) & "«·ﬂ„Ì… «·„ÊÃÊœ…:" & RsTemp("totalqty").value
             '       MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
             '       Exit Sub
             '   End If
            End If
        
        ElseIf Me.GridTrans = ShowPrice Then
            'NO Check Requried
            '    ElseIf (Me.GridTrans = ReturnSalling Or Me.GridTrans = PurchaseTransaction) And Trim(Me.TxtSerial.text) <> "" Then
            '  Set RsTemp = New ADODB.Recordset
            '  Set RsTemp = GetItemQuantityStock(LngItemID, val(Me.storename.BoundText), _
            '      Me.DtpBillDate.value, , , , Trim(Me.TxtSerial.text))
            '  If Not (RsTemp.BOF Or RsTemp.EOF) Then
            '      Msg = "ÌÊÃœ Õ«·Ì« ›Ï «·„Œ“‰ ‰›” —ﬁ„ «·”Ì—Ì«· ·‰›” «·’‰› ...!!"
            '      MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
            '      Me.TxtSerial.SetFocus
            '      Exit Sub
            '  End If
        Else
            'NO Check Requried
        
        End If
    
        '-------------------------------------------------------------------------------
    End If

    With Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
        End If

        If Not Me.TxtSerial Is Nothing Then
            AddItemInGrid .Rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.Text)
            ' customer_screen.bill_no.Caption = frmsalebill.XPTxtBillID.text ' TxtTransSerial.text 'salim2
            'customer_screen.List1.AddItem DcboItemName.BoundText & Chr(9) & DcboItemName.text & "" & Val(TxtQuantity.text) & "*" & Me.TxtPrice.text
            'customer_screen.List2.AddItem Val(TxtQuantity.text) * Me.TxtPrice.text
            '  customer_screen.total
        Else
            AddItemInGrid .Rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1
            '        customer_screen.bill_no.Caption = frmsalebill.XPTxtBillID.text ' TxtTransSerial.text 'salim2
        
            ' customer_screen.bill_no.Caption = Me.TxtTransSerial.text
            '       customer_screen.List1.AddItem DcboItemName.BoundText & Chr(9) & Val(TxtQuantity.text) & " * " & Me.TxtPrice.text  '& Chr(9) & " = " & Val(TxtQuantity.text) * Me.TxtPrice.text
            '        customer_screen.List2.AddItem Val(TxtQuantity.text) * Me.TxtPrice.text
            '        customer_screen.total
        End If

    End With

    CountItems

    If Me.GridTrans <> MoveItems And Me.GridTrans <> OpeningBalance And Me.GridTrans <> Returntransaction Then
        Calculate Me.Grid.Rows - 1
    End If

    If Me.GridTrans = Returntransaction Then
   
        FrmReturnpurchases.Fg_AfterEdit 1, 1
     
    End If

    ClearTopDataEntry
    'Wael Update
    If Not TxtCustCode Is Nothing And Me.GridTrans2 = InvoiceTransactionCompose Then
        TxtCustCode.SetFocus
    Else
        DCboItemCode.SetFocus
    End If
      'Wael Update
    'Me.DCboItemCode.SetFocus
    'Me.DCboItemCode.SetFocus
    'Me.Grid.SetFocus
    Exit Sub
ErrTrap:
End Sub


Private Sub CmdAddSerialLIst_Click()
    Dim LngItemID As Long
    Dim LngStoreID As Long
    LngItemID = val(Me.DCboItemName.BoundText)
    PopUpSerialList LngItemID
    ClearTopDataEntry
End Sub
Function checkitemhaveBarcode(ItemID As Double) As Boolean
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    StrSQL = "select * From TblItems  where  ItemID=" & ItemID
    ' StrSQL = StrSQL & "  barCodeNO='" & DCboItemCode.BoundText & "'"
         Set RsTemp = New ADODB.Recordset
     
     If RsTemp.RecordCount > 0 Then
                If IsNull(RsTemp("barCodeNO").value) Then
                checkitemhaveBarcode = False
                
                Else
               checkitemhaveBarcode = True
                End If
     
  
     Else
     checkitemhaveBarcode = False
     
     End If
     
     
End Function

Private Sub DCboItemCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
TxtPrice.Text = 0
TxtQuantity.Text = 0

 '   If DCboItemCode.BoundText <> "" Then
 '       DcboItemName.BoundText = DCboItemCode.BoundText
 '   Else
 '       Exit Sub
 '   End If


    StrSQL = "select * From TblItems where ItemID=" & DCboItemCode.BoundText
'     StrSQL = StrSQL & " or barCodeNO='" & DCboItemCode.BoundText & "'"
     
     
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
  '      LngItemID = val(DCboItemCode.BoundText)
        LngItemID = IIf(IsNull(RsTemp("itemid").value), 0, RsTemp("ItemID").value)
        If LngUnitID = 0 Then LngUnitID = 1
    GetDefaultItemUnit LngItemID, LngUnitID
    
    
        If LngItemID <> 0 Then
        DCboItemName.BoundText = LngItemID
    Else
        Exit Sub
    End If
    
    
        If Not (IsNull(RsTemp("ItemCase").value)) = True Then
            If RsTemp("ItemCase").value = 1 Then
                Me.CboItemCase.ListIndex = 0
            ElseIf RsTemp("ItemCase").value = 2 Then
                Me.CboItemCase.ListIndex = 1
            End If
        End If

        If RsTemp("HaveSerial").value = True Then
            If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
                             If Not TxtSerial Is Nothing Then
                                 TxtSerial.Enabled = True
                            End If
                If Not Me.CmdAddSerialLIst Is Nothing Then
'                    Me.CmdAddSerialLIst.Enabled = True
                End If
            End If

            '            TxtSerial.Tag = "T"
            TxtQuantity.Enabled = False
            TxtQuantity.Text = "1"
            TxtQuantity.Tag = "T"

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = 3 Then
         
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
                
                                 If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                         TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
                                            If TxtPrice.Text = 0 Then
                                            TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
                                            End If
                        End If
                
                
                
            ElseIf Me.GridTrans = ProductionOrder Then
                TxtPrice.Text = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
                
                
          
                     
           ' ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
          ElseIf GridTrans = ReturnDestruction Or Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                '›Ì Õ«·… «–‰ «·’—›
                If TxtSerial Is Nothing Then
                TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                Else
                 TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, TxtSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                End If
                
        
            Else
                TxtPrice.Text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
       
              If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                 TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
                
                
            End If
        
        ElseIf RsTemp("HaveSerial").value = False Then

            If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice And Me.GridTrans <> ProductionOrder And Me.GridTrans <> 3 Then
                 If Not TxtSerial Is Nothing Then
                 TxtSerial.Enabled = False
                End If
                If Not Me.CmdAddSerialLIst Is Nothing Then
           '         Me.CmdAddSerialLIst.Enabled = False
                End If
            
            End If
        
            TxtQuantity.Tag = "F"
            TxtQuantity.Enabled = True

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Or Me.GridTrans = 28 Then
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
                             If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                 TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
            ElseIf GridTrans = ReturnDestruction Or Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                '›Ì Õ«·… «–‰ «·’—›
                If TxtSerial Is Nothing Then
                TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                Else
                 TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , TxtSerial, SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                End If
                
            ElseIf Me.GridTrans = ProductionOrder Then
                TxtPrice.Text = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
            
            ElseIf Me.GridTrans = Returntransaction Then
                  '  getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            
                'Fromdate = FirstPeriodDateInthisYear
                'GetlastPurchasedata 20, (LngItemID), Fromdate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                
                 TxtPrice.Text = LastPurchasePrice
                 
            Else
        
              '  TxtPrice.Text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
                
            
          TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, 1)   ' ”⁄— «·‘—«¡
                          If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                 TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
                
                
            End If
        End If
  
   
   
    End If

    Exit Sub
ErrTrap:
 
End Sub


Public Sub updateProfit()
On Error Resume Next
    '     If Me.Grid.ColIndex("ItemProfit") = -1 Then Exit Sub
    Dim RowNum As Integer
    Dim LngItemID As Long
    Dim LngUnitID As Long

    With Me.Grid

        For RowNum = 1 To .Rows - 1

            If .TextMatrix(RowNum, .ColIndex("Code")) <> "" Then

                ' FG_AfterEdit RowNum, Fg.ColIndex("Price")
                'Grid_AfterEdit RowNum, .ColIndex("Code")
                If m_GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Or m_GridTrans = Destruction Then
                    LngItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
             '       LngUnitID = val(.TextMatrix(RowNum, .ColIndex("UnitId")))
    LngUnitID = val(Grid.Cell(flexcpData, RowNum, Grid.ColIndex("UnitID")))
   Dim transferid As Double
   
    If m_GridTrans = MoveItems Then
    
  transferid = Get_movingreciveTransaction_ID(val(TxtInvID.Text))
  .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(transferid), LngUnitID, val(Me.StoreName.BoundText))
  Else
  .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
    End If
    
                    
       
       
                End If
                    
                Calculate (RowNum)
                CalculateInvProfit (RowNum)
            End If
    
        Next RowNum

    End With

End Sub

Private Sub DCboItemCode_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF5 Then
 
 FillItemData (val(StoreName.BoundText))
FillGrid val(StoreName.BoundText)

End If


End Sub

Private Sub DCboItemName_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF5 Then
  FillItemData
FillGrid
End If


End Sub

 

Private Sub DCboItemCode_GotFocus()
    DCboItemCode.backcolor = vbYellow
End Sub

Private Sub DCboItemCode_KeyDown(KeyCode As Integer, _
                                 Shift As Integer)
On Error Resume Next
'If DCboItemCode.text = "" Then Exit Sub
     Dim LngItemID As Long
    Dim LngUnitID As Long
'workWithBarcode = True
'
    If KeyCode = vbKeyReturn Then
                    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InventoryOut Then
'
                                    If SystemOptions.workWithBarcode = True Then
                                                                                       LngItemID = val(DCboItemCode.BoundText)
    GetDefaultItemUnit LngItemID, LngUnitID
    
                                               TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
                                               
                                            
                                                If val(DCboItemCode.BoundText) <> 0 And FrmCustomerDisplay.Visible = True Then
                                               
                                                
                                               
                                                
                                            '''''FrmCustomerDisplay.LblInformation.Caption = DcboItemName.text & vbNewLine & "  Price   " & val(Me.TxtPrice.text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
                                           
                                               FrmCustomerDisplay.lblitemid = val(DCboItemCode.BoundText)
                                          

                                              End If
                                             
                                               
                                    SendKeys "{TAB}"
                                    SendKeys "{TAB}"
                                    SendKeys "{TAB}"
                                          
                                          TxtQuantity.Text = 1
                                          CmdAddData_Click
                                          DoEvents
                                        
                                        
                                        
                                        
                                        '  TxtPrice.SetFocus
                                        '   TxtPrice_KeyDown 13, 1

   
                                    DCboItemCode.SetFocus
                                    
'                                                If SystemOptions.AllowItemsShortName = True Then
'
'             TxtShortName.SetFocus
'
'            End If
            
                                            Exit Sub
                                    Else
                                    
'                  If SystemOptions.AllowItemsShortName = True Then
            
'             TxtShortName.SetFocus
'
'            End If
            
                                            GoTo ll
                                    End If
                    
                    End If
End If

ll:




    If KeyCode = vbKeyReturn Then
   
   
        SendKeys "{TAB}"

        
    End If

End Sub

Private Sub DCboItemCode_LostFocus()
    DCboItemCode.backcolor = vbWhite
    'TxtPrice.BackColor = vbWhite
End Sub

'
Private Sub DcboItemName_Change()
     On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
TxtPrice.Text = 0
TxtQuantity.Text = ""
    
    If DCboItemName.BoundText <> "" Then
        DCboItemCode.BoundText = DCboItemName.BoundText
    Else
        Exit Sub
    End If

    LngItemID = val(DCboItemCode.BoundText)
    GetDefaultItemUnit LngItemID, LngUnitID

    StrSQL = "select * From TblItems where ItemID=" & DCboItemName.BoundText
    'customer_screen.Adodc1.RecordSource = StrSQL
    'customer_screen.Adodc1.Refresh
    'MsgBox "salim2"

    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If Not (IsNull(RsTemp("ItemCase").value)) = True Then
            If RsTemp("ItemCase").value = 1 Then
                Me.CboItemCase.ListIndex = 0
            ElseIf RsTemp("ItemCase").value = 2 Then
                Me.CboItemCase.ListIndex = 1
            End If
        End If

        If RsTemp("HaveSerial").value = True Then
            If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
                TxtSerial.Enabled = True

                If Not Me.CmdAddSerialLIst Is Nothing Then
                   ' Me.CmdAddSerialLIst.Enabled = True
                End If
            End If

            TxtQuantity.Enabled = False
            TxtQuantity.Text = "1"

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
             '               ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
             ElseIf GridTrans = ReturnDestruction Or Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
             
                '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                '›Ì Õ«·… «–‰ «·’—›
                If TxtSerial Is Nothing Then
                TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                Else
                 TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, TxtSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                End If
                
            Else
            
                TxtPrice.Text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
            End If

        ElseIf RsTemp("HaveSerial").value = False Then

            If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
                TxtSerial.Enabled = False

                If Not Me.CmdAddSerialLIst Is Nothing Then
                    Me.CmdAddSerialLIst.Enabled = False
                End If
            End If

            TxtQuantity.Enabled = True
            TxtQuantity.Text = ""

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
                
                              If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                 TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
            If TxtPrice.Text = 0 Then
            TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
            End If
            
                End If
                
                
                
            ElseIf GridTrans = ReturnDestruction Or Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                '›Ì Õ«·… «–‰ «·’—›
                If TxtSerial Is Nothing Then
                TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
            Else
            TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, TxtSerial.Text, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
            End If
                          ElseIf Me.GridTrans = Returntransaction Then
                
             
                getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            
                Fromdate = FirstPeriodDateInthisYear
                GetlastPurchasedata 20, (LngItemID), Fromdate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                
                 TxtPrice.Text = LastPurchasePrice
                   
                   
            Else
                TxtPrice.Text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
                
          
                 If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                 TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
                
                
                
            End If
        End If
    End If


 ' Dim LngUnitID As Long

   
   
    Exit Sub
ErrTrap:
End Sub

'
Private Sub DCboItemName_GotFocus()
    DCboItemName.backcolor = vbYellow

    If DCboItemName.BoundText = "" Then
'        SendKeys "{F4}"
    End If

    'TxtPrice.BackColor = vbWhite
End Sub


Private Sub DcboItemName_KeyDown(KeyCode As Integer, _
                                 Shift As Integer)

        If KeyCode = vbKeyReturn Then
   
           SendKeys "{TAB}"
    End If
    
    Dim LngItemID As Long
Dim LngStoreID As Long
If KeyCode = vbKeyReturn Then
  'Dim LngStoreID As Long
  Dim rdate As Date
''/////////
    LngItemID = val(DCboItemName.BoundText)
    If Not TxtLotNo Is Nothing Then
    If ChicIsLotNo(LngItemID) = True Then
                LngStoreID = val(Me.StoreName.BoundText)
                rdate = DtpBillDate.value
                If LngStoreID = 0 Or LngItemID = 0 Then Exit Sub
                ItemLotNoFun 0, 0, LngItemID, LngStoreID, rdate, 1
   End If
   End If
   End If
   
   

End Sub

Private Sub DCboItemName_LostFocus()
    'DCboItemName.BackColor = vbYellow
    DCboItemName.backcolor = vbWhite
End Sub

Private Sub DtpBillDate_DblClick()
updateProfit
End Sub

Private Sub GrdTBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim Msg As String
    Dim IntRes As Integer
   On Error GoTo ErrTrap

    Select Case Button.Key

        Case "AddRow"

            If Me.Grid.TextMatrix(Me.Grid.Rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
                Me.Grid.Rows = Me.Grid.Rows + 1
                Me.GridDefaultValue Me.Grid.Rows - 1
                Me.Grid.Row = Me.Grid.Rows - 1
                Me.Grid.Col = Me.Grid.ColIndex("Code")
                Me.Grid.ShowCell Me.Grid.Rows - 1, Me.Grid.ColIndex("Code")
                Me.Grid.SetFocus
            End If

        Case "RemoveRow"
        Dim F As Form
            For Each F In Forms
        If F.Name = "frmsalebill1" Then
      '      MsgBox "Form Loaded"
            Exit Sub
        End If
    Next F
    
   
            

            If Grid.Rows > 1 Then
                If Grid.Rows = 2 Then
                    Me.Grid.Clear flexClearScrollable, flexClearEverything
                    Me.Calculate 1, True, , True
                Else

                    If Me.Grid.Rows > 1 Then
                        If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                            Me.Grid.RemoveItem (Me.Grid.Row)
                        End If
                    End If
            
SentTypeVAT
                    Calculate 1, , , True
                End If
            End If
            
            
           '           If VatGrid.Rows > 1 Then
           '                                         If VatGrid.Rows = 2 Then
           '                                             Me.VatGrid.Clear flexClearScrollable, flexClearEverything
           '                                         Else
           '
           '                                                             If VatGrid.Rows > 1 Then
           '
           '                                                                     VatGrid.RemoveItem (Grid.Row)
           '
           '                                                             End If
           '
           '                                         End If
           ' End If
        RelinVatGrid
        Case "ClearAll"
'If mdifrmmain.POSTRansactiosG.Visible = True Then

'  frmmangerlogon.show vbModal
'Exit Sub
'End If
            
            Msg = "”Ê› Ì „ Õ–› ﬂ· «·”ÿÊ— «·„ÊÃÊœ… ›Ï «·ÃœÊ· ."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ Â–Â «·ŒÿÊ…...ø"
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

            If IntRes = vbYes Then
                
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Grid.Rows = Me.Grid.FixedRows + 1
                Me.Grid.Row = 1
                Me.VatGrid.Clear flexClearScrollable, flexClearEverything
                Me.VatGrid.Rows = Me.Grid.FixedRows + 1
                Me.VatGrid.Row = 1
                
            End If
RelinVatGrid
        Case "FillData"

            If Me.GridTrans = OpeningBalance Then
                If Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Code")) <> "" Then
                    Grid.Rows = Grid.Rows + 1
                    Me.GridDefaultValue Me.Grid.Rows - 1
                End If

          '      FrmFillBalance.show vbModal
            Else

                If Me.Grid.TextMatrix(Me.Grid.Rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
                    Me.Grid.Rows = Me.Grid.Rows + 1
                    Me.GridDefaultValue Me.Grid.Rows - 1
                End If

             '   FrmFillItems.DealingForm = Me.GridTrans
             '   FrmFillItems.show vbModal
             '   Calculate 1, , , True
            End If

        Case "Tree"

            If Button.value = tbrPressed Then
                Set MyAddItemForm = New FrmAddItem
                Set MyAddItemForm.cGrid = Me
                MyAddItemForm.show
            Else
                CloseMyAddItemForm
            End If

        Case "Barcode"

            If Me.Grid.Parent.XPTxtBillID.Text <> "" Then
                Load FrmPrintBarcode
                FrmPrintBarcode.LBLID.Caption = Me.Grid.Parent.XPTxtBillID.Text
                FrmPrintBarcode.show vbModal
            End If

        Case "GridCols"
        If SystemOptions.hidecolumn = True Then Exit Sub
            ShowGridColFrm
    End Select

    If Me.GridTrans = ProductionOrder Then
        'If FrmProductionOrder.Visible = True Then
        'FrmProductionOrder.show_parts
        'End If
    End If
'salman FillGridValuAdded
    Exit Sub
ErrTrap:
End Sub
Function GetMaxUnitFactors(Optional ItemID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(UnitFactor) AS MaxUnitFactor"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxUnitFactors = IIf(IsNull(rs2("MaxUnitFactor").value), 0, rs2("MaxUnitFactor").value)
Else
GetMaxUnitFactors = 0
End If
End Function
Function GetMaxUnitID(Optional ItemID As Double, Optional UnitFactor As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     UnitID "
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ") and UnitFactor=" & UnitFactor & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxUnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
Else
GetMaxUnitID = 0
End If
End Function
 Public Sub CalculteMaxQtyUnit(LongRow As Long)
Dim UnitFactor As Double
Dim LngItemID As Double
Dim MaxUnitID As Double
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
UnitFactor = GetMaxUnitFactors(LngItemID)
MaxUnitID = GetMaxUnitID(LngItemID, UnitFactor)
Grid.TextMatrix(LongRow, Grid.ColIndex("MaxUnitID")) = MaxUnitID
If UnitFactor <> 0 Then
Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Count")))) / UnitFactor, 3)
Else
Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = 0
End If
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, _
                          ByVal Col As Long)
 
    Dim RowNum As Integer
    Dim SearchResault As Integer
    Dim RsSerial As ADODB.Recordset
    Dim RsFind As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    Dim RsPrice As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim LngItemID As Long
    Dim LngUnitID As Long

    On Error GoTo ErrTrap

    If TxtFillData.Text = "T" Then Exit Sub
    If Row = -1 Or Col = -1 Then Exit Sub
    If RsItems Is Nothing Then Exit Sub

    With Me.Grid

        Select Case .ColKey(Col)

            Case "Code"

                If Grid.TextMatrix(Row, Col) <> "" Then
                    LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    RsItems.find "ItemID=" & LngItemID & "", , adSearchForward, adBookmarkFirst

                    If RsItems.EOF Or RsItems.BOF Then
                        Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Grid.TextMatrix(Row, Col) = ""
                        Me.Grid.Row = Row
                        Me.Grid.Col = Col
                        Exit Sub
                    Else
                        AddItemInGrid Row, LngItemID
                        LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                        LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
               
                        LngUnitID = val(Grid.Cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                
                        '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                        If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Or m_GridTrans = Returntransaction Or m_GridTrans = 28 Then
            
                            AddItemPrice LngItemID, Col, Row, LngUnitID
                            
                                                If SystemOptions.LinkSupplerWithItem = True And m_GridTrans = Returntransaction Then
                         Customer.BoundText = GetSuppler(LngItemID)
                         End If
                         
                        ElseIf m_GridTrans = PurchaseTransaction Then
                            AddItemPrice LngItemID, Col, Row, LngUnitID, 1
                         If SystemOptions.LinkSupplerWithItem = True Then
                         Customer.BoundText = GetSuppler(LngItemID)
                         End If
                          
                        End If
                        
                        
                           
                          If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                 .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
                
            
            
            
            
           '*************************
                     If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                     .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
                                        If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                                        .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID)
                                        End If
                    End If
                
                
                
                
           '**************
                    End If

                Else
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                    'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.Cell(flexcpText, Row, Col, Row, .Cols - 1) = ""
                    Me.Grid.Row = Row
                    Me.Grid.Col = Col
                End If
                   
            Case "Name"

                If Grid.TextMatrix(Row, Col) <> "" Then
                    LngItemID = val(Grid.ComboData)
                    RsItems.find "ItemID=" & LngItemID, , adSearchForward, 1

                    If RsItems.EOF Or RsItems.BOF Then
                        Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Grid.TextMatrix(Row, Col) = ""
                        Me.Grid.Row = Row
                        Me.Grid.Col = Col
                        Exit Sub
                    Else
                 '   If Me.DtpBillDate.value <> Null Then
                        
                        AddItemInGrid Row, LngItemID, , , , , , , , , , , , , , , Me.DtpBillDate.value
                       
                       
                        LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                        LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
               
                        LngUnitID = val(Grid.Cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                
                        '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                        If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Or m_GridTrans = Returntransaction Or m_GridTrans = 28 Then
                            AddItemPrice LngItemID, Col, Row, LngUnitID
                        ElseIf m_GridTrans = PurchaseTransaction Then
                            AddItemPrice LngItemID, Col, Row, LngUnitID, 1
                       
                           
                           
                           
                        End If
                           
                          If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                 .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName)
                End If
                
                
           '*************************
                 If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                 .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
                 If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                 .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID)
                 End If
                End If
                 '*************************
                                
                                
            
                    End If

                Else
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                    'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.TextMatrix(Row, Col) = ""
                    Me.Grid.Row = Row
                    Me.Grid.Col = Col
                    Exit Sub
                End If



Dim LngStoreID As Long
    Dim rdate As Date
''//////////
    LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
                
              If ChicIsLotNo(LngItemID) = True Then
                
                LngStoreID = val(Me.StoreName.BoundText)
                rdate = DtpBillDate.value
           '     ItemLotNoFun Row, Col, LngItemID, LngStoreID, rdate, 0
                ItemLotNoFun Row, Col, LngItemID, val(Grid.TextMatrix(Row, Grid.ColIndex("StoreID2"))), rdate, 0
           
           
           End If
           
              '     Case "Width", "Height"
             'Grid.TextMatrix(Row, Grid.ColIndex("Area")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Height")))
             
             
             
                        
            Case "Width", "Height", "NoCount", "length", "OUTR", "INR"
            
            'SalimHereAnd Wael
                checkQtybylengthoption Row, LngItemID
            'If Codition
            
                  
             
              
            Case "Area"
             Grid.TextMatrix(Row, Grid.ColIndex("Width")) = 0
             Grid.TextMatrix(Row, Grid.ColIndex("Height")) = 0
             Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
             Case "NoCount"
             Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
             
 
           
             Case "Serial"
            
                LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                SearchResault = Grid.FindRow(Grid.TextMatrix(Row, Grid.ColIndex("Serial")), , Grid.ColIndex("Serial"))

 If Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                    '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
              '      StrSerial = Trim$(.TextMatrix(Row, .ColIndex("Serial")))
                    Dim UnitID As Long
 
                    UnitID = val(.Cell(flexcpData, Row, .ColIndex("UnitID")))
                
                    .TextMatrix(Row, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, Grid.TextMatrix(Row, Grid.ColIndex("Serial")), , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), UnitID, val(Me.StoreName.BoundText))
               
     End If
               
                If SearchResault <> Row And SearchResault > 0 Then
                    StrSQL = "Select * From TblItems where ItemID=" & Trim(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    Set RsSerial = New ADODB.Recordset
                    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    StrSQL = "Select * From TblItems where ItemID=" & Trim(Grid.TextMatrix(SearchResault, Grid.ColIndex("Code")))
                    Set RsFind = New ADODB.Recordset
                    RsFind.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Msg = " „ «œŒ«· «·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                    Msg = Msg + RsSerial("ItemName").value & CHR(13)
                    Msg = Msg + "ﬁ»· –·ﬂ «·Ì «·’‰›" & CHR(13)
                    Msg = Msg + RsFind("ItemName").value & CHR(13)
                    Msg = Msg + " √ﬂœ „‰ «·”Ì—Ì«· «·Œ«’ »«·√’‰«›"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    .TextMatrix(Row, Col) = ""
                End If

                '(N)«· √ﬂœ √‰ «·ﬁÿ⁄Â „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ ›Ì «·Õ«·…
                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = Destruction Or Me.GridTrans = MoveItems Then
                
                    If SystemOptions.SysAllowStockNegative = False Then
                        If TxtModFlag.Text = "N" Then
                            If .TextMatrix(Row, .ColIndex("Code")) <> "" Then
                                LngItemID = val(.TextMatrix(Row, .ColIndex("Code")))
                                StrSQL = "Select * From TblItems where ItemID=" & LngItemID
                                Set RsTemp = New ADODB.Recordset
                                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                                If RsTemp("HaveSerial").value = True Then
                                    If Trim(.TextMatrix(Row, .ColIndex("Serial"))) <> "" Then
                                        If StoreName.BoundText <> "" Then
                                            Set RsTest = New ADODB.Recordset
                                            Set RsTest = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), , , True, , Trim(.TextMatrix(Row, .ColIndex("Serial"))))

                                            If RsTest.EOF Or RsTemp.BOF Then
                                                Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· : "
                                                Msg = Msg + Trim(.TextMatrix(Row, .ColIndex("Serial"))) & CHR(13) & "„‰ «·’‰› : " & Trim(.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))) & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & Row
                                                Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & CHR(13)
                                                Msg = Msg + "Â·  —€»  ›Ì „⁄—›… ”Ì—Ì«· «·ﬁÿ⁄ «·„ÊÃÊœ… ›Ï «·„Œ«“‰"
                                
                                                If (MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)) = vbYes Then
                                                    FrmSearchSerial.Tag = RsTemp("ItemCode").value

                                                    If Me.GridTrans = InvoiceTransaction Then
                                                        FrmSearchSerial.LngGridRow = Row
                                                        FrmSearchSerial.Txt.Text = "Search"
                                                    ElseIf Me.GridTrans = Returntransaction Then
                                                        FrmSearchSerial.Txt.Text = "Retutn"
                                                    ElseIf Me.GridTrans = MoveItems Then
                                                        FrmSearchSerial.Txt.Text = "Moving"
                                                    ElseIf Me.GridTrans = Destruction Then
                                                        FrmSearchSerial.Txt.Text = "Destruction"
                                                    End If

                                                    'FrmSearchSerial.Cmd_Click 0
                                                    FrmSearchSerial.show vbModal
                                                End If

                                                'Exit Sub
                                            End If
                                        End If

                                        RsTest.Close

                                        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Then
                                      'salman      CalculateInvProfit Row
                                        End If
                                        
                                        
                                    End If
                                End If
                            End If

                            RsTemp.Close
                        End If
                    End If
                End If

            Case "guaranteeTime" ' «·÷„«‰

                If .TextMatrix(Row, .ColIndex("guaranteeTime")) <> "" Then
                    If Not (IsNumeric(.TextMatrix(Row, .ColIndex("guaranteeTime")))) Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) < 1 Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) > 36 Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) <> Int(val(.TextMatrix(Row, .ColIndex("guaranteeTime")))) Then
                        Msg = "Â–Â «·ﬁÌ„… €Ì— ’«·Õ…" & CHR(13)
                        Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì… ’ÕÌÕ… »Ì‰ 1 -36 " & CHR(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    End If
                End If

            Case "Count"

              '  If mdifrmmain.ActiveForm.name = "FrmReturnSalling" Or mdifrmmain.ActiveForm.name = "FrmReturnpurchases" Then
             If m_GridTrans = ReturnSalling Or m_GridTrans = Returntransaction Then
                    If m_GridTrans = ReturnSalling Then
                        StrSQL = "   SELECT     SUM(ShowQty) AS ShowQty FROM         dbo.Transaction_Details   Where Transaction_ID=" & val(FrmReturnSalling.TxtInvID.Text) & " and item_id=" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                  '
  
                    Else
                    
                    
                        StrSQL = " SELECT     SUM(ShowQty) AS ShowQty  FROM         dbo.Transaction_Details where  Transaction_ID=" & val(FrmReturnpurchases.TxtInvID.Text) & " and item_id=" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    End If

                    Dim rs As New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        If rs("showqty") < val(Grid.TextMatrix(Row, Grid.ColIndex("count"))) Then MsgBox "«·ﬂ„Ì… «·„»«⁄… „‰ Â–« «·’‰› «ﬁ· „‰ «·ﬂ„Ì… «·„–ﬂÊ—… ... ÌÃ» „—«Ã⁄… «·›« Ê—…": Exit Sub
                    End If

                    rs.Close
                End If

                LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))

                '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Then
                    '      AddItemPrice LngItemID, Col, Row, LngUnitID
                StrSQL = "Select FreeQty From TblItems Where ItemID=" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                
                  Dim Rs1 As New ADODB.Recordset
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                      If Not (Rs1.BOF Or Rs1.EOF) Then
                                      If Not IsNull(Rs1("FreeQty").value) Then
                                      
                                                    If Rs1("FreeQty").value > 0 Then
                                                    Grid.TextMatrix(Row + 1, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) * Rs1("FreeQty").value / 100
                                                    End If
                                        
                                    
                                    End If
                    
                    Rs1.Close

                 End If
            End If
            
                 If Me.GridTrans = ShipmentRegisteration And .ColIndex("RecivedShippedQty1") <> -1 Then
            .TextMatrix(Row, .ColIndex("RecivedShippedQty1")) = .TextMatrix(Row, .ColIndex("Count"))
            End If
            
         '***************************************************
        If Me.GridTrans = ShipmentRegisteration And .ColIndex("RecivedShippedQty1") <> -1 Then
            .TextMatrix(Row, .ColIndex("RecivedShippedQty1")) = .TextMatrix(Row, .ColIndex("Count"))
         End If
               If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = MoveItems Then
                                If SystemOptions.AllowExperDateFIFO = True Then
                         '        ItemLotNoFunFIFo Row, val(.TextMatrix(Row, .ColIndex("Code"))), val(.TextMatrix(Row, .ColIndex("Count"))), val(.TextMatrix(Row, .ColIndex("Price"))), , , , val(.TextMatrix(Row, .ColIndex("UnitID"))), , , , , , , , , , , val(StoreName.BoundText), GetQty(LngItemID, Row)
                                  ItemLotNoFunFIFo Row, val(.TextMatrix(Row, .ColIndex("Code"))), val(.TextMatrix(Row, .ColIndex("Count"))), val(.TextMatrix(Row, .ColIndex("Price"))), , , , val(.TextMatrix(Row, .ColIndex("UnitID"))), , , , , , , , , , , val(.TextMatrix(Row, .ColIndex("StoreID2"))), GetQty(LngItemID, Row)
                                 End If
             End If
                
             '***************************************************
             
             
             
           Case "ScurrencyID"

              '  If mdifrmmain.ActiveForm.name = "FrmReturnSalling" Or mdifrmmain.ActiveForm.name = "FrmReturnpurchases" Then
              
                     
                        StrSQL = "Select * From currency Where  id=" & val(Grid.TextMatrix(Row, Grid.ColIndex("ScurrencyID")))
                  

             
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                      Grid.TextMatrix(Row, Grid.ColIndex("Scurrenyrate")) = IIf(IsNull(rs("rate").value), 1, rs("rate").value)
                    
                    End If

                    rs.Close
            
 
            
            Case "UnitID"
            
                LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
          LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
   '          LngUnitID = val(Grid.Cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                   
                   
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Then
                If SystemOptions.IsGeometricProportions Then
                    
                    LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
                    If LngUnitID = 0 Then
                        GetDefaultItemUnit LngItemID, LngUnitID
                    End If
                     StrSQL = "SELECT     TOP 100 PERCENT PartItemQty,ForUnit,MethodCalc "
                    
                    StrSQL = StrSQL + " FROM         dbo.TblItemsUnits"
                    
                    StrSQL = StrSQL + " Where (dbo.TblItemsUnits.ItemID = " & LngItemID & ")"
                    StrSQL = StrSQL + " and (dbo.TblItemsUnits.UnitID = " & LngUnitID & ")"
                    Dim RsParts As ADODB.Recordset
                    Set RsParts = New ADODB.Recordset
                    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
                    If Not (RsParts.EOF Or RsParts.BOF) Then
                        Grid.TextMatrix(Row, Grid.ColIndex("ItemQty")) = RsParts!PartItemQty & ""
                        Grid.TextMatrix(Row, Grid.ColIndex("ForUnit")) = RsParts!ForUnit & ""
                        Grid.TextMatrix(Row, Grid.ColIndex("MethodCalc")) = RsParts!MethodCalc & ""
                    End If
                    checkQtybylengthoption Row, LngItemID
                End If
                
            End If
                   
              Me.Grid.Cell(flexcpData, Row, Col) = Grid.ComboData

                '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Or m_GridTrans = 28 Then
                    AddItemPrice LngItemID, Col, Row, LngUnitID
                
                    Me.Grid.Cell(flexcpData, Row, Col) = Grid.ComboData
                    PutDefCount Row
                ElseIf m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Then
                    LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))

                    AddItemPrice LngItemID, Col, Row, LngUnitID
                
                ElseIf m_GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                  
                    .TextMatrix(Row, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DtpBillDate.value, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                ElseIf m_GridTrans = PurchaseTransaction Then
                    AddItemPrice LngItemID, Col, Row, LngUnitID, 1
            
                End If

                'Case "UnitID"
                '            LngItemID = Val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                '            LngUnitID = Val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
                '
                '            '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                '            If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Then
                '                AddItemPrice LngItemID, Col, Row, LngUnitID
                '            End If
                '            Case "UnitID"
                '            Me.Grid.Cell(flexcpData, Row, Col) = Grid.ComboData
                '            PutDefCount Row
                '            If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Then
                '            LngItemID = Val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                '            LngUnitID = Val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
                '
                '                AddItemPrice LngItemID, Col, Row, LngUnitID

                   '*******************************************19032017*********
                .TextMatrix(Row, .ColIndex("Remarks")) = ""
                  .TextMatrix(Row, .ColIndex("QtyFaqtors")) = 0
           .TextMatrix(Row, .ColIndex("QtyFaqtors")) = DescUnitFact(LngItemID, LngUnitID)
           
          If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(Row, .ColIndex("Remarks")) = .TextMatrix(Row, .ColIndex("Remarks")) & "  ﬂÊ‰ „‰"
            .TextMatrix(Row, .ColIndex("Remarks")) = .TextMatrix(Row, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, LngUnitID)
            Else
            '.TextMatrix(Row, .ColIndex("Remarks")) = .TextMatrix(Row, .ColIndex("Remarks")) & "Consisting of"
            '.TextMatrix(Row, .ColIndex("Remarks")) = .TextMatrix(Row, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, LngUnitID)
            End If
           
                    '*******************************************19032017*********
        
                   Dim LngCatlog As Long
            Case "catalog"
            
                LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                LngCatlog = val(Grid.TextMatrix(Row, Grid.ColIndex("catalog")))
                Me.Grid.Cell(flexcpData, Row, Col) = Grid.ComboData
              Case "TotalPrice"
              If val(Grid.TextMatrix(Row, Grid.ColIndex("TotalPrice"))) <> 0 Then
              If val(Grid.TextMatrix(Row, Grid.ColIndex("Vatyo"))) <> 0 Then
                Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(Grid.TextMatrix(Row, Grid.ColIndex("TotalPrice"))) / ((val(Grid.TextMatrix(Row, Grid.ColIndex("Vatyo"))) / 100) + 1)
                Else
                Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(Grid.TextMatrix(Row, Grid.ColIndex("TotalPrice")))
                End If
              End If
                
     Case "NoHours"
             Grid.TextMatrix(Row, Grid.ColIndex("TotalPriceNoHours")) = val(Grid.TextMatrix(Row, Grid.ColIndex("PriceNoHours"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("NoHours")))
           
             
      Case "MixNo"
    Dim StoreId2 As Double
        If SystemOptions.RawMaterMix = True Then
         StoreId2 = GetStoreID2(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), Grid.TextMatrix(Row, Grid.ColIndex("MixNo")))
         If StoreId2 <> 0 Then
                         Grid.TextMatrix(Row, Grid.ColIndex("StoreID2")) = StoreId2
         End If
        End If
        
        
        End Select

    End With

    If m_GridTrans <> Returntransaction And m_GridTrans <> OpeningBalance And Me.GridTrans <> NewGard Then
        If val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 0 Or val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 1 Or val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 Then
            Grid.TextMatrix(Row, Grid.ColIndex("DiscountVal")) = 0
        End If
        
           If Grid.ColIndex("ExtraType") <> -1 Then
                If val(Grid.TextMatrix(Row, Grid.ColIndex("ExtraType"))) = 0 Or val(Grid.TextMatrix(Row, Grid.ColIndex("ExtraType"))) = 1 Then
            Grid.TextMatrix(Row, Grid.ColIndex("ExtraVal")) = 0
        End If
        
        End If
    End If

    With Grid

        Select Case .ColKey(Col)

            Case "UnitID", "Name", "Code"

              'salman    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ProductionOrder Then
               'salman       Me.Calculate Row, , , True
               'salman   ElseIf Me.GridTrans = PurchaseTransaction Then
                 'salman     Me.Calculate Row
               'salman   Else
                'salman      Me.Calculate Row
               'salman   End If
           Case "Count", "Price", "DiscountType", "DiscountVal"
                'salman    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ProductionOrder Then
                'salman      Me.Calculate Row, , , True
                'salman  ElseIf Me.GridTrans = PurchaseTransaction Then
                  'salman    Me.Calculate Row
               'salman   Else
                  'salman    Me.Calculate Row
               'salman   End If
               

 Case "L", "W", "H1", "H2"
      If Me.GridTrans = InvoiceTransaction Then
             Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("L"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("W"))) * ((val(Grid.TextMatrix(Row, Grid.ColIndex("H1"))) + val(Grid.TextMatrix(Row, Grid.ColIndex("H2")))) / 2)
             Grid.TextMatrix(Row, Grid.ColIndex("Count")) = Grid.TextMatrix(Row, Grid.ColIndex("Count")) / 1000000
      End If
        
             
        End Select

    End With

    CountItems
    HideComment

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = 28 Then
         Me.Calculate Row, , , True
      '  Me.Calculate Row
    ElseIf Me.GridTrans = 19 Or Me.GridTrans = 14 Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = Template Or Me.GridTrans = Destruction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = OpeningBalance Then
        Me.Calculate Row
               Else
                    Me.Calculate Row
                  End If
               
  

    '-------------------------
    'Me.Grid.AutoSize 0, Grid.Cols - 1, False
    '-------------------------
    With Grid

        If .TextMatrix(Row, .ColIndex("Code")) <> "" Then
            EditGrid Row, Col
            '.Row = .Row - 1
        Else
        
            If .Rows = .FixedRows + 1 Then
                .Clear flexClearScrollable, flexClearEverything
                Me.GridDefaultValue 1
            Else
                .RemoveItem Row
            End If
        End If

    End With

    With Grid

        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Then
            CalculateInvProfit Row
        End If



Grid.TextMatrix(Row, Grid.ColIndex("SalesValue")) = GetItemPrice(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), , val(Grid.Cell(flexcpData, Row, Grid.ColIndex("UnitID"))))
   CalculteMaxQtyUnit Row
   
    End With
    ReCalVAT Row, 1
'salman TxtDiscount_Val_Change
    Exit Sub
ErrTrap:
End Sub

 
Private Sub checkQtybylengthoption(ByVal Row As Long, ByVal LngItemID As Long)
    
     If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Then
        Dim mWidth As Double, mHeight As Double, mNoCount  As Double, mLength As Double, lowering As Double, mFlgX As Double, ShowQty As Double
        Dim mIsCountIsGeometricProportions As Boolean
        Dim ForUnit  As Double, MethodCalc As Double, Qty As Double, mQtyGrid As Double, OUTR As Double, INR As Double
        
        With Grid
           If SystemOptions.IsGeometricProportions Then
                     
                ForUnit = val(.TextMatrix(Row, .ColIndex("ForUnit")))
                mQtyGrid = val(.TextMatrix(Row, .ColIndex("NoCount")))
                mWidth = val(.TextMatrix(Row, .ColIndex("Width")))
                mHeight = val(.TextMatrix(Row, .ColIndex("Height")))
                mNoCount = val(.TextMatrix(Row, .ColIndex("NoCount")))
                mLength = val(.TextMatrix(Row, .ColIndex("length")))
                OUTR = val(.TextMatrix(Row, .ColIndex("OUTR")))
                INR = val(.TextMatrix(Row, .ColIndex("INR")))
                ShowQty = val(.TextMatrix(Row, .ColIndex("Count")))
                
                
                
                
                MethodCalc = val(.TextMatrix(Row, .ColIndex("MethodCalc")))
                'lowering = val(.TextMatrix(Row, .ColIndex("lowering")))
                
                     '  increase = val(.TextMatrix(i, .ColIndex("increase")))
                Qty = val(.TextMatrix(Row, .ColIndex("ItemQty")))
                If val(ForUnit) = 0 Then ForUnit = 1
                If Trim(ForUnit) = "" And MethodCalc = 0 Then Exit Sub
                    
                If MethodCalc = 1 Then 'ﬂ„Ì…
                    mFlgX = Round(Qty, 2)
                ElseIf MethodCalc = 2 Then '⁄—÷
                    Qty = ((val(mWidth) / ForUnit) * Qty) - lowering
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                ElseIf MethodCalc = 3 Then 'ÿÊ·
                    Qty = ((val(mHeight) / ForUnit) * Qty) - lowering
                    
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                    
                ElseIf MethodCalc = 4 Then 'ÿÊ·+⁄—÷
                    Qty = ((val(mWidth) + val(mHeight)) / ForUnit * Qty) - lowering
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                ElseIf MethodCalc = 5 Then 'ÿÊ·*⁄—÷
                    Qty = ((val(mWidth) * val(mHeight)) / ForUnit * Qty) - lowering
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                ElseIf MethodCalc = 6 Then ' «·ÿÊ· ·ﬂ· ⁄—÷
                    Qty = ((val(mHeight) / ForUnit) - lowering) * Qty * val(mWidth)
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                ElseIf MethodCalc = 7 Then ' «·⁄—÷ ·ﬂ· ÿÊ·
                    Qty = ((val(mWidth) / ForUnit * Qty) * val(mHeight)) - lowering   ' ((val(mWidth) +  / ForUnit * Qty) - lowering
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                    
                ElseIf MethodCalc = 8 Then ' «·«— ›«⁄ * «·⁄—÷ * ÿÊ·
                    Qty = (((val(mWidth) * val(mHeight) * val(mLength))) / ForUnit * Qty) - lowering
                    mFlgX = Round(Qty, 2)
                    .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 2)
                                          
                ElseIf MethodCalc = 9 Then 'ÿÊ·+⁄—÷
                            
                    Qty = (val(mLength) * 3.14 * ((val(IIf(OUTR = 0, INR, OUTR)) / 2) ^ 2) / ForUnit * Qty) - lowering
                               
                ElseIf MethodCalc = 10 Then 'ÿÊ·+⁄—÷
                    Qty = (((val(mWidth) * val(mLength) * val(mHeight))) / ForUnit * Qty) - lowering
                ElseIf MethodCalc = 11 Then 'ÿÊ·+⁄—÷
                '    Qty = (val(mLength) * 3.14 * ((val(OUTR * OUTR) - val(INR * INR))) / ForUnit * Qty)
                         Qty = (val(mLength) * 3.14 * (val(OUTR / 2 * OUTR / 2) - val(INR / 2 * INR / 2))) / ForUnit * Qty
                End If
                If MethodCalc <> 1 Then
                    If val(.TextMatrix(Row, .ColIndex("Count"))) = 0 Then
                         mFlgX = Round(Qty, 2)
                        .TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 3)
               
                  Else
                    mFlgX = Qty
'                    .TextMatrix(Row, .ColIndex("Count")) = mFlgX * val(mQtyGrid)
.TextMatrix(Row, .ColIndex("Count")) = Round(mFlgX * val(mQtyGrid), 3)
                    '
                    End If
                End If
               
            End If
            End With
        End If
            If mFlgX * val(mQtyGrid) <> 0 Then mIsCountIsGeometricProportions = True Else mIsCountIsGeometricProportions = False
                   
               
             '  End If
               '  ÕœÌœ «·„⁄«œ·… ÿ»ﬁ« ··’‰› „‰ „·› «·’‰› ‰›”Â
               ' ÿÊ· Ê⁄—÷ Ê«— ›«⁄ Êﬁÿ— Œ«—ÃÌ Êﬁÿ— œ«Œ·Â
               'XQty⁄‰ ÿ—ÌﬁÂ«  Õ’· ⁄·Ì ﬂ„Ì… „»œ√Ì…
            'NoCount    ÷—»  ›Ì «·⁄œœ
            'showqty=NoCount*XQty
            
              With Grid
               Dim widthPrice As Double
                Grid.TextMatrix(Row, Grid.ColIndex("Area")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Height")))
                 
                     LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
                   If SystemOptions.AllowPriceWithWidth = True Then
                              widthPrice = GetItemPriceByWitdth(LngItemID, val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))))
                              .TextMatrix(Row, .ColIndex("Price")) = widthPrice * val(Grid.TextMatrix(Row, Grid.ColIndex("Width")))
                   Else
                           
                          If mIsCountIsGeometricProportions Then
                               Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
                           End If
                   End If
            End With
             
End Sub

 
 
 Sub ReCalVAT(Optional Row As Long, Optional flg As Integer)
 'salimsalim 22 07 2019
 If SystemOptions.AllItemInVAT = True Then
 SentTypeVAT
Exit Sub
 End If
 
 If Me.GridTrans = InvoiceTransaction Then
CalculteValueAdded Row, 21, flg
End If
 If Me.GridTrans = PurchaseTransaction Then
CalculteValueAdded Row, 22
End If

     If Me.GridTrans = ReturnSalling Then
      CalculteValueAdded Row, 9
     End If
      If Me.GridTrans = Returntransaction Then
         CalculteValueAdded Row, 5
      End If
        If Me.GridTrans = purchaseOrderApproved Then
         CalculteValueAdded Row, 29
      End If
      If Me.GridTrans = salespricelist Then
         CalculteValueAdded Row, 42
      End If
      If Me.GridTrans = ShowPrice Then
         CalculteValueAdded Row, 30
      End If
End Sub
Public Function GetSuppler(Optional ItemID As Long) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select DefaultSupplier from TblItems where ItemID=" & ItemID & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetSuppler = IIf(IsNull(rs2("DefaultSupplier").value), 0, rs2("DefaultSupplier").value)
Else
GetSuppler = 0
End If
End Function

   

 Public Sub CalculteValueAdded(LongRow As Long, Optional TransType As Integer, Optional flg As Integer = 0, Optional AllItems As Integer = 0, Optional posDelete As Boolean = False)
If SystemOptions.PriceWithVAT = True And (TransType = 21 Or TransType = 9) Then Exit Sub
 If (Me.TxtModFlag.Text = "R" Or Me.TxtModFlag.Text = "" Or val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code"))) = 0) And posDelete = False Then Exit Sub
 Dim Percentg As Double
Dim LngItemID As Double
If StoreName Is Nothing Then Exit Sub
If mdifrmmain.taxes.Visible = True Then
'If TransType = 9 And ReturnSales = True Then
If ReturnTyp = 2 Then
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
Percentg = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")))
Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Round((Percentg * (val(val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) - CalculteDiscountLin(LongRow)))) / 100, 2)
If AllItems = 0 Then
FillGridValuAdded TransType, LongRow
End If
Else
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
If SystemOptions.AllItemInVAT = True Then

Percentg = val(cCompanyInfo.VATItems)
Else
Percentg = PercentgValueAdded(DtpBillDate, val(StoreName.BoundText), LngItemID, TransType)
End If
If Percentg = -1 Then
Percentg = 0
If SystemOptions.UserInterface = ArabicInterface Then
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "„⁄›Ì"
Else
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "Exempt"
End If
Else
If Grid.ColIndex("TypeVAT") <> -1 Then 'salim1503
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = Percentg
End If

End If
If Grid.ColIndex("Vatyo") <> -1 Then
Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")) = Percentg
End If
If Grid.ColIndex("Vat") <> -1 Then
 If val(CboDiscount_Type.ListIndex) > 0 And val(TxtDiscount_Val.Text) <> 0 Then
 If TransType <> 21 Then
    Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu")))) - CalculteDiscountLin(LongRow)))) / 100, ".##")
 Else
    
    Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = (Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu")))) - CalculteDiscountLin(LongRow)))) / 100

   
End If

            If Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = "." Then
             Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = 0
            End If

Else
Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".##")
If Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = "." Then
 Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = 0
End If

End If
End If
If AllItems = 0 Then
FillGridValuAdded TransType, LongRow, flg
End If
End If
End If
End Sub
  


Public Sub CalculteValueAddedxxx15042019(LongRow As Long, Optional TransType As Integer, Optional flg As Integer = 0, Optional AllItems As Integer = 0, Optional posDelete As Boolean = False)
If SystemOptions.PriceWithVAT = True And (TransType = 21 Or TransType = 9) Then Exit Sub
 If (Me.TxtModFlag.Text = "R" Or Me.TxtModFlag.Text = "" Or val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code"))) = 0) And posDelete = False Then Exit Sub
 Dim Percentg As Double
Dim LngItemID As Double
If StoreName Is Nothing Then Exit Sub
If mdifrmmain.taxes.Visible = True Then
'If TransType = 9 And ReturnSales = True Then
If ReturnTyp = 2 Then
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
Percentg = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")))
Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Round((Percentg * (val(val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) - CalculteDiscountLin(LongRow)))) / 100, 2)
If AllItems = 0 Then
FillGridValuAdded TransType, LongRow
End If
Else
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
If SystemOptions.AllItemInVAT = True Then

Percentg = val(cCompanyInfo.VATItems)
Else
Percentg = PercentgValueAdded(DtpBillDate, val(StoreName.BoundText), LngItemID, TransType)
End If
If Percentg = -1 Then
Percentg = 0
If SystemOptions.UserInterface = ArabicInterface Then
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "„⁄›Ì"
Else
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "Exempt"
End If
Else
If Grid.ColIndex("TypeVAT") <> -1 Then 'salim1503
Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = Percentg
End If

End If
If Grid.ColIndex("Vatyo") <> -1 Then
Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")) = Percentg
End If
If Grid.ColIndex("Vat") <> -1 Then
 If val(CboDiscount_Type.ListIndex) > 0 And val(TxtDiscount_Val.Text) <> 0 Then
Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu")))) - CalculteDiscountLin(LongRow)))) / 100, ".##")
            If Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = "." Then
             Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = 0
            End If

Else
Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".##")
If Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = "." Then
 Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = 0
End If

End If
End If
If AllItems = 0 Then
FillGridValuAdded TransType, LongRow, flg
End If
End If
End If
End Sub
  
  Public Function CalculteDiscountLin(Optional i As Long)
Dim DblDiscountTotal As Double
Dim TotalDiscountPerLine As Double
Dim Total As Double
If LBLGross Is Nothing Then Exit Function
If LBLGross = "" Then
LBLGross = 0
End If
With Grid
     If .Rows > 1 Then
            Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
     End If
           If Not CboDiscount_Type Is Nothing Then
             If CboDiscount_Type.ListIndex = 1 Then
                  DblDiscountTotal = val(TxtDiscount_Val.Text)
               Else
                                                
              If CboDiscount_Type.ListIndex = 2 Then
               If TxtDiscount_Val.Text <> "" Then
                DblDiscountTotal = ((val(TxtDiscount_Val.Text) * LBLGross) / 100)
               End If
              End If
             End If
           End If
     ''////////////
          If val(.TextMatrix(i, .ColIndex("Valu"))) > 0 And LBLGross > 0 Then
           TotalDiscountPerLine = .TextMatrix(i, .ColIndex("Valu")) / (LBLGross) * DblDiscountTotal
           TotalDiscountPerLine = Round(TotalDiscountPerLine, 20)
           Else
           TotalDiscountPerLine = 0
           End If
         
End With
CalculteDiscountLin = TotalDiscountPerLine
End Function
Public Sub FillGridValuAddedAllItems()
Dim i As Integer
If VatGrid Is Nothing Then Exit Sub
With VatGrid
.Clear flexClearScrollable, flexClearEverything
.Rows = Grid.Rows
For i = 1 To Grid.Rows - 1
If Grid.ColIndex("Vat") <> -1 Then 'salim1503
If val(Grid.TextMatrix(i, Grid.ColIndex("Vat"))) <> 0 And val(Grid.TextMatrix(i, Grid.ColIndex("Code"))) <> 0 Then
End If
.TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("Code")) = Grid.Cell(flexcpTextDisplay, i, Grid.ColIndex("Code"))
.TextMatrix(i, .ColIndex("Name")) = Grid.Cell(flexcpTextDisplay, i, Grid.ColIndex("Name"))
.TextMatrix(i, .ColIndex("select")) = True
.TextMatrix(i, .ColIndex("Vat")) = Grid.TextMatrix(i, Grid.ColIndex("Vat"))
.TextMatrix(i, .ColIndex("Vatyo")) = Grid.TextMatrix(i, Grid.ColIndex("Vatyo"))
.TextMatrix(i, .ColIndex("ItemID")) = Grid.TextMatrix(i, Grid.ColIndex("Code"))
.TextMatrix(i, .ColIndex("Valu")) = Grid.TextMatrix(i, Grid.ColIndex("Valu"))
End If
Next i
End With
RelinVatGrid

End Sub
Public Sub FillGridValuAdded(Optional TransType As Integer, Optional LongRow As Long, Optional flg As Integer = 0)
Dim i As Integer
Dim cunt As Integer
Dim Name As String
Dim code As String
Dim Ros As Long
Dim start As Integer
cunt = 0
If VatGrid Is Nothing Then Exit Sub
With VatGrid
.Rows = Grid.Rows
 Ros = LongRow
start = LongRow
For i = start To Ros
If Grid.ColIndex("Vat") <> -1 Then 'salim1503
cunt = start
'RetriveItems val(Grid.TextMatrix(i, Grid.ColIndex("Code"))), Name, code
.TextMatrix(cunt, .ColIndex("index")) = cunt
.TextMatrix(cunt, .ColIndex("Code")) = Grid.Cell(flexcpTextDisplay, i, Grid.ColIndex("Code"))
.TextMatrix(cunt, .ColIndex("Name")) = Grid.Cell(flexcpTextDisplay, i, Grid.ColIndex("Name"))
.TextMatrix(cunt, .ColIndex("select")) = True
.TextMatrix(cunt, .ColIndex("Vat")) = Grid.TextMatrix(i, Grid.ColIndex("Vat"))
.TextMatrix(cunt, .ColIndex("Vatyo")) = Grid.TextMatrix(i, Grid.ColIndex("Vatyo"))
.TextMatrix(cunt, .ColIndex("ItemID")) = Grid.TextMatrix(i, Grid.ColIndex("Code"))
.TextMatrix(cunt, .ColIndex("Valu")) = Grid.TextMatrix(i, Grid.ColIndex("Valu"))
End If
Next i
End With
RelinVatGrid
End Sub



Public Sub RetriveItems(Optional ItemID As Double, Optional ByRef Name As String, Optional ByRef code As String)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " Select * from TblItems where ItemID=" & ItemID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
code = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
If SystemOptions.UserInterface = ArabicInterface Then
Name = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
Name = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
Else
Name = ""
code = ""
End If
End Sub
Private Sub EditGrid(Row As Long, _
                     Col As Long)
    Exit Sub

    Dim i As Long

    With Me.Grid

        'On Error Resume Next
        If m_EditDir = EditDirVer Then
            If .Row + 1 < .Rows Then
                .Row = Row + 1
            ElseIf .Row + 1 = .Rows Then
                .Rows = .Rows + 1
                .Row = Row + 1
            Else
                .Row = .Row - 1
            End If

            .Col = Col
            .ShowCell .Row, .Col
            GridDefaultValue .Row
            .EditCell
        ElseIf m_EditDir = EditDirHor Then
            .Row = Row

            For i = Col + 1 To .Cols - 1

                If .ColHidden(i) = False Then
                    .Col = i
                    Exit For
                End If

            Next i

            .ShowCell .Row, .Col
            GridDefaultValue .Row
            .EditCell
        End If

    End With

End Sub

Private Sub Grid_AfterSort(ByVal Col As Long, _
                           Order As Integer)

    With Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
        End If

    End With

End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, _
                                 ByVal Col As Long)
    On Error Resume Next

    With Me.Grid

        If .ColIndex("AddSerials") <> -1 Then
            .ColWidth(.ColIndex("AddSerials")) = 350
        End If

    End With

End Sub
  Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    On Error GoTo ErrTrap

    With Me.Grid

        Select Case .ColKey(Col)
        
  
        
    Case "Valu"
    Cancel = True
    Exit Sub
Case "ShipedQty"
Cancel = True

Case "ISSUEDQTY"
Cancel = True

Case "TotalInvoiceQty"
Cancel = True


Case "RecivedShippedQty"
Cancel = True
Case "TotallRecivedShippedQty"
Cancel = True

            Case "AddSerials"

                If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexUnchecked Then
                    Cancel = True
                End If

            Case "Name"

                If .ColComboList(.ColIndex("Name")) = "" Then
                    Cancel = True
                End If

            Case "HaveSerial"
                Cancel = True

            Case "Count"

                If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexChecked Then
                    Cancel = True
                Else
                    Cancel = False
                End If


            Case "Serial", "guaranteeTime"

                If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexChecked Then
                    Cancel = False
                Else
                    Cancel = True
                End If

        End Select

    End With

    If m_GridTrans <> Returntransaction And m_GridTrans <> OpeningBalance Then
        If val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 0 Or val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 1 Or val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 Then
            Grid.TextMatrix(Row, Grid.ColIndex("DiscountVal")) = 0

            If Col = Grid.ColIndex("DiscountVal") Then
                Cancel = True
            End If
        End If

        If Col = Grid.ColIndex("Valu") Then
            Cancel = True
        End If
    End If
    
    
If m_GridTrans = InvoiceTransaction Then
If Not CBoBasedON Is Nothing Then
With Grid
If val(CBoBasedON.ListIndex) = 4 Then
If SystemOptions.AllowChangeSalesAtTransfer = False Then
If .ColKey(Col) = "DiscountType" Or .ColKey(Col) = "DiscountVal" Then
Cancel = False
Else
Cancel = True
End If
Else
Cancel = False
End If
Else
Cancel = False
End If

Select Case .ColKey(Col)
Case "StoreID2"

                If .TextMatrix(Row, .ColIndex("MixNo")) <> "" Then
                    Cancel = True
                Else
                    Cancel = False
                End If
                
End Select

End With
End If
End If

ErrTrap:
End Sub
 

 
 
 Public Function Calculate(LngRow As Long, _
                          Optional BolOneRow As Boolean = False, _
                          Optional BolSavae As Boolean = False, _
                          Optional TaxCalc As Boolean = False) As Boolean

    Dim RowNum As Integer
    Dim Msg As String
    Dim Total As Double
    Dim RealTotal As Double
    Dim TotalNet As Double
    Dim DblDiscountTotal As Double
    Dim DblRowTotal As Double
        Dim DblRowTotal2 As Double
      Dim DblRowTotal3 As Double
           
        Dim DblRowTotals As Double
        
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
    Dim StrSQL  As String
    
       Dim DblCommissionTotal As Double
               Dim Commisionvalue As Double
       
'    On Error GoTo ErrTrap

    CountItems
    If Not LBLGross Is Nothing Then
LBLGross = 0
End If
    With Grid
        ReSerialGrid
    LblTotalQty = 0
    
If Not lblTotalSalesPrice Is Nothing Then lblTotalSalesPrice = 0

        For RowNum = .FixedRows To .Rows - 1
            
           ' .TextMatrix(LngRow, .ColIndex("UnitPrice")) = .TextMatrix(LngRow, .ColIndex("Price"))
'            If SystemOptions.IsGeometricProportions = True And val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) <> 0 Then
'                If val(.TextMatrix(LngRow, .ColIndex("Count"))) <> 0 Then
'                    .TextMatrix(LngRow, .ColIndex("Price")) = val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) / val(.TextMatrix(LngRow, .ColIndex("Count")))
'                End If
'            End If

            If .ColIndex("UnitPrice") <> -1 Then
                        If SystemOptions.IsGeometricProportions = True And val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) <> 0 Then
                        If val(.TextMatrix(LngRow, .ColIndex("Count"))) <> 0 Then
                            .TextMatrix(LngRow, .ColIndex("Price")) = val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) * val(.TextMatrix(LngRow, .ColIndex("NoCount"))) / val(.TextMatrix(LngRow, .ColIndex("Count")))
                            End If
                        End If
            End If

            .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                    
                    
             If .ColIndex("showPrice1") <> -1 And .ColIndex("ShowQty1") <> -1 Then
                        If .ColIndex("ScurrencyID") <> -1 Then
                        DblRowTotal2 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice1"))) * val(.TextMatrix(RowNum, .ColIndex("Scurrenyrate"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty1"))), SystemOptions.SysDefCurrencyForamt)
                        Else
                        DblRowTotal2 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice1"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty1"))), SystemOptions.SysDefCurrencyForamt)
                        End If
                
                End If


   If .ColIndex("Value1") <> -1 Then
.TextMatrix(RowNum, .ColIndex("Value1")) = DblRowTotal2
End If

      If .ColIndex("showPrice2") <> -1 And .ColIndex("ShowQty2") <> -1 Then
                DblRowTotal3 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice2"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty2"))), SystemOptions.SysDefCurrencyForamt)
  End If

   If .ColIndex("Value2") <> -1 Then
.TextMatrix(RowNum, .ColIndex("Value2")) = DblRowTotal3
End If


   If .ColIndex("TotalCost") <> -1 And .ColIndex("Salaries1") <> -1 And .ColIndex("Salaries2") <> -1 Then
.TextMatrix(RowNum, .ColIndex("TotalCost")) = val(.TextMatrix(RowNum, .ColIndex("Value1"))) _
+ val(.TextMatrix(RowNum, .ColIndex("Value2"))) _
+ val(.TextMatrix(RowNum, .ColIndex("Salaries1"))) + val(.TextMatrix(RowNum, .ColIndex("Salaries2")))
           '     If val(.TextMatrix(RowNum, .ColIndex("Price"))) = 0 Then
           If val(.TextMatrix(RowNum, .ColIndex("TotalCost"))) <> 0 Then
                 .TextMatrix(RowNum, .ColIndex("Price")) = .TextMatrix(RowNum, .ColIndex("TotalCost"))
          End If
                 
                 .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
            '    End If
                
                

End If



            LngCurItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
            If Grid.ColIndex("UnitID") <> -1 Then 'salim1503

            LngUnitID = val(.Cell(flexcpData, RowNum, .ColIndex("UnitID")))    'val(.TextMatrix(RowNum, .ColIndex("UnitID")))
End If
        '    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            
        '    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            If Me.GridTrans = NewGard Then
            
            Else
            
          StrSQL = "SELECT     dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsUnits.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitFactor, dbo.TblItemsUnits.DefaultUnit, "
 StrSQL = StrSQL + "                       dbo.TblItemsUnits.UnitSalesPrice, dbo.TblItemsUnits.UnitPurPrice, dbo.TblItemsUnits.FactorByDefaultUnit, dbo.TblItemsUnits.FactorBySmallUnit,"
 StrSQL = StrSQL + "                       dbo.TblItemsUnits.SecOrder , dbo.TblItemsUnits.JunckID"
 StrSQL = StrSQL + " FROM         dbo.TblItemsUnits INNER JOIN"
 StrSQL = StrSQL + "                       dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
               StrSQL = StrSQL + "  where  ItemID=" & LngCurItemID
            
            StrSQL = StrSQL + " AND TblItemsUnits.UnitID=" & LngUnitID
            
            Dim RsUnitData As New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                 
                LblTotalQty = LblTotalQty + (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(RsUnitData("UnitFactor").value))
                     If SystemOptions.UserInterface = ArabicInterface Then
                     
                        .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitName").value
             '            .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitName").value
                    Else
                    If IsNull(RsUnitData("UnitNamee").value) Then
             '   .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitName").value
                                .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitName").value
                     Else
             '       .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitNamee").value
                       .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitNamee").value
                     End If
                     
                    End If
                   If .ColIndex("SalesPrice") <> -1 Then
                        .TextMatrix(RowNum, .ColIndex("SalesPrice")) = IIf(IsNull(RsUnitData("UnitSalesPrice").value), 0, RsUnitData("UnitSalesPrice").value)
                        .TextMatrix(RowNum, .ColIndex("TotalSalesPrice")) = Round(val(RsUnitData!UnitSalesPrice & "") * val(.TextMatrix(RowNum, .ColIndex("Count"))), 2)
                        lblTotalSalesPrice = val(lblTotalSalesPrice) + val(.TextMatrix(RowNum, .ColIndex("TotalSalesPrice")))
                   End If
            End If

            
            RsUnitData.Close
        
        End If
      If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
        With Me.Grid
          .TextMatrix(RowNum, .ColIndex("Valu")) = Int(.TextMatrix(RowNum, .ColIndex("Valu")))
        End With
End If

    
        Next RowNum

        If .Rows > 1 Then
            Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
            RealTotal = Total
        End If

        If Not Me.LblTotalAll Is Nothing Then
            Me.LblTotalAll.Caption = Round(Total, SystemOptions.SysDefCurrencyForamt)
            '          Me.LblTotalAll.Caption = Val(Format(total, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
        End If
    
        DblDiscountTotal = 0
  DblCommissionTotal = 0
  
        Dim discountvalue As Double
        Dim DblPrice As Double

        If .ColIndex("DiscountType") <> -1 Then

            For RowNum = 1 To .Rows - 1
                DblRowTotal = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                
                DblPrice = val(.TextMatrix(RowNum, .ColIndex("Price")))
          
                Select Case val(.TextMatrix(RowNum, .ColIndex("DiscountType")))

                                Case 0
                                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                                    
                                    
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 1
                                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                                    
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 2
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) - val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                             '     DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                    DblDiscountTotal = DblDiscountTotal + val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                                    discountvalue = val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
            
                                Case 3
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) / 100))
                                    DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                    DblDiscountTotal = DblDiscountTotal + ((val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
                                    discountvalue = ((val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) * DblRowTotal) / 100)
            
                                Case 4
                              
                                DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                      .TextMatrix(RowNum, .ColIndex("Valu")) = 0
                                    DblDiscountTotal = DblDiscountTotal + DblRowTotal
                                    discountvalue = DblRowTotal
                            End Select

 .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal - discountvalue
'**********************************************************************
        If .ColIndex("ExtraType") <> -1 Then
Select Case val(.TextMatrix(RowNum, .ColIndex("ExtraType")))

                                Case 0
                               
                                    Commisionvalue = 0
            
                                Case 1
                                
                                    Commisionvalue = 0
            
                                Case 2
                                 '   .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) + val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
                                    DblCommissionTotal = DblCommissionTotal + val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
                                    Commisionvalue = val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
            
                                Case 3
                                 '   .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) / 100))
                                    DblCommissionTotal = DblCommissionTotal + ((val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
                                    Commisionvalue = ((val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
            
                         
                            End Select
                           .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal + Commisionvalue - discountvalue
            End If
'*************************************************************************************


                         If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Then
                                            If .ColIndex("discountvalue") <> -1 Then
                                                .TextMatrix(RowNum, .ColIndex("discountvalue")) = discountvalue
                                     
                                            End If
                                            
                                               If .ColIndex("Commisionvalue") <> -1 Then
                                                .TextMatrix(RowNum, .ColIndex("Commisionvalue")) = Commisionvalue
                                     
                                            End If
                                            
                        
                         End If
             
             
       If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
        With Me.Grid
          .TextMatrix(RowNum, .ColIndex("Valu")) = Int(.TextMatrix(RowNum, .ColIndex("Valu")))
        End With
End If
If Not LBLGross Is Nothing Then
LBLGross = LBLGross + val(.TextMatrix(RowNum, .ColIndex("Valu")))
End If
            Next RowNum
If .Rows > 1 Then
        Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
End If
                        If Me.GridTrans <> ShowPrice And Me.GridTrans <> Template And Me.GridTrans <> Destruction And Me.GridTrans <> Returntransaction And Me.GridTrans <> StockSettlement Then
            
                            'Õ”«» ﬁÌ„… «·Œ’„
                                                 If Not CboDiscount_Type Is Nothing Then
                                                                    If CboDiscount_Type.ListIndex = 1 Or CboDiscount_Type.ListIndex = 3 Then
                                                                    
                                                                        DblDiscountTotal = DblDiscountTotal + val(TxtDiscount_Val.Text)
                                                                     
                                                                        Total = RealTotal - DblDiscountTotal
                                                                    Else
                                                
                                                                                    If CboDiscount_Type.ListIndex = 2 Then
                                                                                                If TxtDiscount_Val.Text <> "" Then
                                                                                                    DblDiscountTotal = DblDiscountTotal + ((val(TxtDiscount_Val.Text) * Total) / 100)
                                                                                                
                                                                                                    '  total = RealTotal * Val(1 - (TxtDiscount_Val.text / 100))
                                                                                                    Total = RealTotal - DblDiscountTotal
                                                                                                End If
                                                                                    End If
                                                                    End If
                                                End If
                                                
                                                
                                 '**********************************************
                                 
                                 '*************************************************
                                                
                        End If
        End If

        If Not Me.LblDiscountsTotal Is Nothing Then
            Me.LblDiscountsTotal.Caption = (DblDiscountTotal)
        End If

        If Not Me.LblCommision Is Nothing Then
            Me.LblCommision.Caption = (DblCommissionTotal)
        End If
        
    End With

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Then
        '   CalculateInvProfit LngRow
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Then
        TotalNet = Total

        '÷—»Ì… Œ’„ Ê≈÷«›…
        If Not Me.TxtAddTax Is Nothing Then
            If val(Me.TxtAddTax.Text) > 0 Then
                Dim DblTemp As Double

                If Not Me.LblTaxAddValue Is Nothing Then
                    LblTaxAddValue.Caption = ((val(Me.TxtAddTax.Text) * Total) / 100)
                End If

                TotalNet = Total * (1 - val((Me.TxtAddTax.Text)) / 100)
            End If
        End If

        '÷—»Ì… œ„€…
        If Not Me.TxtStampTax Is Nothing Then
            If val(Me.TxtStampTax.Text) > 0 Then
                If Not Me.LblTaxStampValue Is Nothing Then
                    Me.LblTaxStampValue.Caption = val(Me.TxtStampTax.Text)
                End If

                TotalNet = TotalNet - val(Me.TxtStampTax.Text)
            End If
        End If

        '÷—»Ì… «·„»Ì⁄« 
        If TaxCalc = True Then
            If TxtTaxValue.Text <> "" Then
                If val(TxtTaxValue.Text) <> 0 Then
                    If Not Me.LblTaxSalesValue Is Nothing Then
                        LblTaxSalesValue.Caption = ((val(Me.TxtTaxValue.Text) * Total) / 100)
                    End If

                    TotalNet = TotalNet + ((val(Me.TxtTaxValue.Text) * Total) / 100)
                End If
            End If
        End If

        '÷—»Ì… Œœ„…
        If Not Me.TxtServiceTax Is Nothing Then
            If val(Me.TxtServiceTax.Text) <> 0 Then
                If Not Me.LblTaxServiceValue Is Nothing Then
                    Me.LblTaxServiceValue.Caption = ((val(Me.TxtServiceTax.Text) * Total) / 100)
                End If

                TotalNet = TotalNet + ((val(Me.TxtServiceTax.Text) * Total) / 100)
            End If
        End If

        TotalNet = Round(TotalNet, SystemOptions.SysDefCurrencyForamt)
        txtTotal.Text = TotalNet
    Else
        txtTotal.Text = Round(Total, SystemOptions.SysDefCurrencyForamt)
    End If

    If Not BolSavae = False Then
        If TxtModFlag.Text = "N" Or TxtModFlag.Text = "E" Then
            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) > val(Total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
                Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Calculate = False
                Exit Function
            End If

            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) < val(Total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
                Msg = Msg + "√ﬁ· „‰ ≈Ã„«·Ì «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Calculate = False
                Exit Function
            End If
        End If
    End If
    

    Calculate = True

    Exit Function
ErrTrap:
End Function



 
 Public Function Calculate13112018(LngRow As Long, _
                          Optional BolOneRow As Boolean = False, _
                          Optional BolSavae As Boolean = False, _
                          Optional TaxCalc As Boolean = False) As Boolean

    Dim RowNum As Integer
    Dim Msg As String
    Dim Total As Double
    Dim RealTotal As Double
    Dim TotalNet As Double
    Dim DblDiscountTotal As Double
    Dim DblRowTotal As Double
        Dim DblRowTotal2 As Double
      Dim DblRowTotal3 As Double
           
        Dim DblRowTotals As Double
        
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
    Dim StrSQL  As String
    
       Dim DblCommissionTotal As Double
               Dim Commisionvalue As Double
       
'    On Error GoTo ErrTrap

    CountItems
    If Not LBLGross Is Nothing Then
LBLGross = 0
End If
    With Grid
        ReSerialGrid
    LblTotalQty = 0

        For RowNum = .FixedRows To .Rows - 1
            .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                    
                    
             If .ColIndex("showPrice1") <> -1 And .ColIndex("ShowQty1") <> -1 Then
                        If .ColIndex("ScurrencyID") <> -1 Then
                        DblRowTotal2 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice1"))) * val(.TextMatrix(RowNum, .ColIndex("Scurrenyrate"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty1"))), SystemOptions.SysDefCurrencyForamt)
                        Else
                        DblRowTotal2 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice1"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty1"))), SystemOptions.SysDefCurrencyForamt)
                        End If
                
                End If


   If .ColIndex("Value1") <> -1 Then
.TextMatrix(RowNum, .ColIndex("Value1")) = DblRowTotal2
End If

      If .ColIndex("showPrice2") <> -1 And .ColIndex("ShowQty2") <> -1 Then
                DblRowTotal3 = Round(val(.TextMatrix(RowNum, .ColIndex("showPrice2"))) * val(.TextMatrix(RowNum, .ColIndex("ShowQty2"))), SystemOptions.SysDefCurrencyForamt)
  End If

   If .ColIndex("Value2") <> -1 Then
.TextMatrix(RowNum, .ColIndex("Value2")) = DblRowTotal3
End If


   If .ColIndex("TotalCost") <> -1 And .ColIndex("Salaries1") <> -1 And .ColIndex("Salaries2") <> -1 Then
.TextMatrix(RowNum, .ColIndex("TotalCost")) = val(.TextMatrix(RowNum, .ColIndex("Value1"))) _
+ val(.TextMatrix(RowNum, .ColIndex("Value2"))) _
+ val(.TextMatrix(RowNum, .ColIndex("Salaries1"))) + val(.TextMatrix(RowNum, .ColIndex("Salaries2")))
           '     If val(.TextMatrix(RowNum, .ColIndex("Price"))) = 0 Then
           If val(.TextMatrix(RowNum, .ColIndex("TotalCost"))) <> 0 Then
                 .TextMatrix(RowNum, .ColIndex("Price")) = .TextMatrix(RowNum, .ColIndex("TotalCost"))
          End If
                 
                 .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
            '    End If
                
                

End If



            LngCurItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
            If Grid.ColIndex("UnitID") <> -1 Then 'salim1503

            LngUnitID = val(.Cell(flexcpData, RowNum, .ColIndex("UnitID")))    'val(.TextMatrix(RowNum, .ColIndex("UnitID")))
End If
        '    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            
        '    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            If Me.GridTrans = NewGard Then
            
            Else
            
          StrSQL = "SELECT     dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItemsUnits.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitFactor, dbo.TblItemsUnits.DefaultUnit, "
 StrSQL = StrSQL + "                       dbo.TblItemsUnits.UnitSalesPrice, dbo.TblItemsUnits.UnitPurPrice, dbo.TblItemsUnits.FactorByDefaultUnit, dbo.TblItemsUnits.FactorBySmallUnit,"
 StrSQL = StrSQL + "                       dbo.TblItemsUnits.SecOrder , dbo.TblItemsUnits.JunckID"
 StrSQL = StrSQL + " FROM         dbo.TblItemsUnits INNER JOIN"
 StrSQL = StrSQL + "                       dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
               StrSQL = StrSQL + "  where  ItemID=" & LngCurItemID
            
            StrSQL = StrSQL + " AND TblItemsUnits.UnitID=" & LngUnitID
            
            Dim RsUnitData As New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                 
                LblTotalQty = LblTotalQty + (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(RsUnitData("UnitFactor").value))
                     If SystemOptions.UserInterface = ArabicInterface Then
                     
                        .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitName").value
             '            .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitName").value
                    Else
                    If IsNull(RsUnitData("UnitNamee").value) Then
             '   .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitName").value
                                .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitName").value
                     Else
             '       .Cell(flexcpData, RowNum, .ColIndex("UnitID")) = RsUnitData("UnitNamee").value
                       .TextMatrix(RowNum, .ColIndex("unitid")) = RsUnitData("UnitNamee").value
                     End If
                     
                    End If
            End If

            
            RsUnitData.Close
        
        End If
      If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
        With Me.Grid
          .TextMatrix(RowNum, .ColIndex("Valu")) = Int(.TextMatrix(RowNum, .ColIndex("Valu")))
        End With
End If

    
        Next RowNum

        If .Rows > 1 Then
            Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
            RealTotal = Total
        End If

        If Not Me.LblTotalAll Is Nothing Then
            Me.LblTotalAll.Caption = Round(Total, SystemOptions.SysDefCurrencyForamt)
            '          Me.LblTotalAll.Caption = Val(Format(total, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
        End If
    
        DblDiscountTotal = 0
  DblCommissionTotal = 0
  
        Dim discountvalue As Double
        Dim DblPrice As Double

        If .ColIndex("DiscountType") <> -1 Then

            For RowNum = 1 To .Rows - 1
                DblRowTotal = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                
                DblPrice = val(.TextMatrix(RowNum, .ColIndex("Price")))
          
                Select Case val(.TextMatrix(RowNum, .ColIndex("DiscountType")))

                                Case 0
                                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                                    
                                    
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 1
                                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                                    
                                    DblDiscountTotal = DblDiscountTotal + 0
                                    discountvalue = 0
            
                                Case 2
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) - val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                             '     DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                    DblDiscountTotal = DblDiscountTotal + val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                                    discountvalue = val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
            
                                Case 3
'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) / 100))
                                    DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                    DblDiscountTotal = DblDiscountTotal + ((val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
                                    discountvalue = ((val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) * DblRowTotal) / 100)
            
                                Case 4
                              
                                DblRowTotal = .TextMatrix(RowNum, .ColIndex("Valu"))
                                      .TextMatrix(RowNum, .ColIndex("Valu")) = 0
                                    DblDiscountTotal = DblDiscountTotal + DblRowTotal
                                    discountvalue = DblRowTotal
                            End Select

 .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal - discountvalue
'**********************************************************************
        If .ColIndex("ExtraType") <> -1 Then
Select Case val(.TextMatrix(RowNum, .ColIndex("ExtraType")))

                                Case 0
                               
                                    Commisionvalue = 0
            
                                Case 1
                                
                                    Commisionvalue = 0
            
                                Case 2
                                 '   .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) + val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
                                    DblCommissionTotal = DblCommissionTotal + val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
                                    Commisionvalue = val(.TextMatrix(RowNum, .ColIndex("ExtraVal")))
            
                                Case 3
                                 '   .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) / 100))
                                    DblCommissionTotal = DblCommissionTotal + ((val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
                                    Commisionvalue = ((val(.TextMatrix(RowNum, .ColIndex("ExtraVal"))) * val(.TextMatrix(RowNum, .ColIndex("Valu")))) / 100)
            
                         
                            End Select
                           .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal + Commisionvalue - discountvalue
            End If
'*************************************************************************************


                         If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Then
                                            If .ColIndex("discountvalue") <> -1 Then
                                                .TextMatrix(RowNum, .ColIndex("discountvalue")) = discountvalue
                                     
                                            End If
                                            
                                               If .ColIndex("Commisionvalue") <> -1 Then
                                                .TextMatrix(RowNum, .ColIndex("Commisionvalue")) = Commisionvalue
                                     
                                            End If
                                            
                        
                         End If
             
             
       If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
        With Me.Grid
          .TextMatrix(RowNum, .ColIndex("Valu")) = Int(.TextMatrix(RowNum, .ColIndex("Valu")))
        End With
End If
If Not LBLGross Is Nothing Then
LBLGross = LBLGross + val(.TextMatrix(RowNum, .ColIndex("Valu")))
End If
            Next RowNum
If .Rows > 1 Then
        Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
End If
                        If Me.GridTrans <> ShowPrice And Me.GridTrans <> Template And Me.GridTrans <> Destruction And Me.GridTrans <> Returntransaction And Me.GridTrans <> StockSettlement Then
            
                            'Õ”«» ﬁÌ„… «·Œ’„
                                                 If Not CboDiscount_Type Is Nothing Then
                                                                    If CboDiscount_Type.ListIndex = 1 Or CboDiscount_Type.ListIndex = 3 Then
                                                                    
                                                                        DblDiscountTotal = DblDiscountTotal + val(TxtDiscount_Val.Text)
                                                                     
                                                                        Total = RealTotal - DblDiscountTotal
                                                                    Else
                                                
                                                                                    If CboDiscount_Type.ListIndex = 2 Then
                                                                                                If TxtDiscount_Val.Text <> "" Then
                                                                                                    DblDiscountTotal = DblDiscountTotal + ((val(TxtDiscount_Val.Text) * Total) / 100)
                                                                                                
                                                                                                    '  total = RealTotal * Val(1 - (TxtDiscount_Val.text / 100))
                                                                                                    Total = RealTotal - DblDiscountTotal
                                                                                                End If
                                                                                    End If
                                                                    End If
                                                End If
                                                
                                                
                                 '**********************************************
                                 
                                 '*************************************************
                                                
                        End If
        End If

        If Not Me.LblDiscountsTotal Is Nothing Then
            Me.LblDiscountsTotal.Caption = (DblDiscountTotal)
        End If

        If Not Me.LblCommision Is Nothing Then
            Me.LblCommision.Caption = (DblCommissionTotal)
        End If
        
    End With

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Then
        '   CalculateInvProfit LngRow
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Then
        TotalNet = Total

        '÷—»Ì… Œ’„ Ê≈÷«›…
        If Not Me.TxtAddTax Is Nothing Then
            If val(Me.TxtAddTax.Text) > 0 Then
                Dim DblTemp As Double

                If Not Me.LblTaxAddValue Is Nothing Then
                    LblTaxAddValue.Caption = ((val(Me.TxtAddTax.Text) * Total) / 100)
                End If

                TotalNet = Total * (1 - val((Me.TxtAddTax.Text)) / 100)
            End If
        End If

        '÷—»Ì… œ„€…
        If Not Me.TxtStampTax Is Nothing Then
            If val(Me.TxtStampTax.Text) > 0 Then
                If Not Me.LblTaxStampValue Is Nothing Then
                    Me.LblTaxStampValue.Caption = val(Me.TxtStampTax.Text)
                End If

                TotalNet = TotalNet - val(Me.TxtStampTax.Text)
            End If
        End If

        '÷—»Ì… «·„»Ì⁄« 
        If TaxCalc = True Then
            If TxtTaxValue.Text <> "" Then
                If val(TxtTaxValue.Text) <> 0 Then
                    If Not Me.LblTaxSalesValue Is Nothing Then
                        LblTaxSalesValue.Caption = ((val(Me.TxtTaxValue.Text) * Total) / 100)
                    End If

                    TotalNet = TotalNet + ((val(Me.TxtTaxValue.Text) * Total) / 100)
                End If
            End If
        End If

        '÷—»Ì… Œœ„…
        If Not Me.TxtServiceTax Is Nothing Then
            If val(Me.TxtServiceTax.Text) <> 0 Then
                If Not Me.LblTaxServiceValue Is Nothing Then
                    Me.LblTaxServiceValue.Caption = ((val(Me.TxtServiceTax.Text) * Total) / 100)
                End If

                TotalNet = TotalNet + ((val(Me.TxtServiceTax.Text) * Total) / 100)
            End If
        End If

        TotalNet = Round(TotalNet, SystemOptions.SysDefCurrencyForamt)
        txtTotal.Text = TotalNet
    Else
        txtTotal.Text = Round(Total, SystemOptions.SysDefCurrencyForamt)
    End If

    If Not BolSavae = False Then
        If TxtModFlag.Text = "N" Or TxtModFlag.Text = "E" Then
            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) > val(Total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
                Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         '       Calculate = False
                Exit Function
            End If

            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) < val(Total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
                Msg = Msg + "√ﬁ· „‰ ≈Ã„«·Ì «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
              '  Calculate = False
                Exit Function
            End If
        End If
    End If
    

'    Calculate = True

    Exit Function
ErrTrap:
End Function



Public Property Get GridTrans() As GridTransType
    GridTrans = m_GridTrans
End Property
 
 Public Property Let GridTrans(ByVal vNewValue As GridTransType)

    If (vNewValue <> ReturnDestruction _
And vNewValue <> purchaserequest And vNewValue <> InventoryOut And vNewValue <> InvoiceTransaction And vNewValue <> PurchaseTransaction And vNewValue <> Returntransaction And vNewValue <> GridTransType.ShowPrice And vNewValue <> Template And vNewValue <> Destruction And vNewValue <> ReturnSalling And vNewValue <> MoveItems And vNewValue <> OpeningBalance And vNewValue <> StockSettlement And vNewValue <> ProductionOrder And vNewValue <> NewGard And vNewValue <> ProductionMaterialReciveVoucher And vNewValue <> purchaseorderrequest And vNewValue <> internalorder And vNewValue <> INVENTORYIN And vNewValue <> salespricelistRequest And vNewValue <> salespricelist And vNewValue <> purchaseorder And vNewValue <> ShipmentOrder And vNewValue <> ShipmentRegisteration And vNewValue <> ShipmentRecieveVoucher And vNewValue <> purchaseOrderApproved And vNewValue <> RecervieProductionVoucherNew And vNewValue <> BookInventories And vNewValue <> NewGard And vNewValue <> internalissuerequesT) Then
    Else
        m_GridTrans = vNewValue
    End If

End Property


Public Sub GridDefaultValue(Crow As Long)
   On Error GoTo ErrTrap

    With Grid
    If .ColIndex("ItemCase") <> -1 Then
        .TextMatrix(Crow, .ColIndex("ItemCase")) = 1
End If
If .ColIndex("Count") <> -1 Then
        If .TextMatrix(Crow, .ColIndex("Count")) = "" Then
            .TextMatrix(Crow, .ColIndex("Count")) = 1
        End If
    
End If
If .ColIndex("DiscountType") <> -1 Then
        If .ColIndex("DiscountType") <> -1 Then
            .TextMatrix(Crow, .ColIndex("DiscountType")) = 1
        End If
 End If
    
    End With
  If SystemOptions.WorkWithLINKEDiTEMS = True Then
  
 '   FillGrid
End If




    Exit Sub
ErrTrap:
End Sub


 Public Sub FillStoredLin()
Dim i As Integer
     If Me.TxtModFlag = "N" Or Me.TxtModFlag = "E" Then
    With Grid
                            For i = 1 To .Rows - 1
                                            If val(.TextMatrix(i, .ColIndex("Code"))) <> 0 Then
                                            If (.TextMatrix(i, .ColIndex("MixNo")) = "" Or val(.TextMatrix(i, .ColIndex("StoreID2"))) = 0) Then
                                                                .TextMatrix(i, .ColIndex("StoreID2")) = StoreName.BoundText
                                                End If

                                            End If
                            Next i
End With

End If
End Sub

Public Function CaseStoredLin() As Integer
If SystemOptions.MultyStore = False Then
CaseStoredLin = 3
Exit Function
End If
Dim i As Integer
Dim CaseNo As Integer
Dim StoreId As Integer
CaseNo = 0

With Grid
StoreId = val(.TextMatrix(1, .ColIndex("StoreID2")))
For i = 1 To .Rows - 1
        If val(.TextMatrix(i, .ColIndex("Code"))) <> 0 And (val(.TextMatrix(i, .ColIndex("StoreID2")))) <> StoreId Then
        StoreId = val(.TextMatrix(i, .ColIndex("StoreID2")))
                        If CaseNo < 2 Then
                        CaseNo = CaseNo + 1
                        End If
        End If
Next i
If StoreId <> val(StoreName.BoundText) And CaseNo = 0 Then
CaseNo = 1
Else
CaseNo = 2
End If
End With
CaseStoredLin = CaseNo
End Function

  
  
 
 Public Sub FillGrid(Optional storename1 As Integer = -1, Optional ItemByCus As Integer = -1, Optional ByVal isCashCust As Boolean = True)
firstrun = True
    Dim StrList As String
    Dim RsNote As ADODB.Recordset
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
   On Error GoTo ErrTrap

Set cCompanyInfo = New ClsCompanyInfo
    m_ShowQtyToolTip = True
If storename1 <> -1 Or ItemByCus = 1 Then GoTo xxx

    With Grid

        If SystemOptions.UserInterface = EnglishInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
            .TextMatrix(0, .ColIndex("Code")) = "Item Code"
            .TextMatrix(0, .ColIndex("Name")) = "Item Name"

      If .ColIndex("NProductionOrderNO") <> -1 Then
                .TextMatrix(0, .ColIndex("NProductionOrderNO")) = "Pro. Order No"
            End If
            
            
            If .ColIndex("Serial") <> -1 Then
                .TextMatrix(0, .ColIndex("Serial")) = "Item Serial"
            End If

            If .ColIndex("ItemCase") <> -1 Then
                .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case"
            Else
        
            End If

            If .ColIndex("HaveSerial") <> -1 Then
                .TextMatrix(0, .ColIndex("HaveSerial")) = "Have SerialS"
            Else
                .TextMatrix(0, .ColIndex("HaveSerial")) = "Have SerialS"

            End If
        
            '.TextMatrix(0, .ColIndex("UnitID")) = "Unit"
         
            .TextMatrix(0, .ColIndex("Count")) = "Quantity"
            .TextMatrix(0, .ColIndex("Price")) = "Price"

            If .ColIndex("Valu") <> -1 Then
                .TextMatrix(0, .ColIndex("Valu")) = "Item Value"
            End If

            If .ColIndex("DiscountVal") <> -1 Then
                .TextMatrix(0, .ColIndex("DiscountVal")) = "Discount Value"
            End If

            If .ColIndex("DiscountType") <> -1 Then
                .TextMatrix(0, .ColIndex("DiscountType")) = "Discount Type"
            End If

            If .ColIndex("guaranteeTime") <> -1 Then
                .TextMatrix(0, .ColIndex("guaranteeTime")) = "Guarantee Interval"
            End If
        
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
            .ColComboList(.ColIndex("ItemCase")) = "#1;ÃœÌœ|#2;„” ⁄„·"
        Else
            .ColComboList(.ColIndex("ItemCase")) = "#1;New|#2;Used"
        End If


        '--------------------------------------------------------------------------
        If m_GridTrans <> OpeningBalance And m_GridTrans <> NewGard Then
            If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("DiscountType")) = "#1;·«ÌÊÃœ Œ’„|#2;Œ’„ »ﬁÌ„…|#3;Œ’„ »‰”»…|#4;„Ã«‰Ì"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("DiscountType")) = "#1;NO Discount|#2;Value Discount|#3;Percent Discount|#4;Free"
            End If
           If .ColIndex("ExtraType") <> -1 Then
        
                  If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("ExtraType")) = "#1;·«ÌÊÃœ ⁄„Ê·…|#2; ⁄„Ê·… ﬁÌ„…|#3;⁄„Ê·…  ‰”»… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("ExtraType")) = "#1;NA|#2;Value commision|#3;Percent commision "
            End If
           End If
            
        
        End If
    
        .ColAlignment(.ColIndex("Serial")) = flexAlignLeftCenter
        '--------------------------------------------
        .Cols = .Cols + 1
        .ColWidthMin = 350
        .ColKey(.Cols - 1) = "AddSerials"
        .ColComboList(.ColIndex("AddSerials")) = "..."
    
        .ColPosition(.ColIndex("AddSerials")) = 5 ' .ColPosition(.ColIndex("Serail")/
    
        '--------------------------------------------
        '≈÷«›… ⁄„Êœ ‰Ê⁄ «·’‰› ≈› —«÷Ì« ›Ï ﬂ· «·‘«‘« 
        If .ColIndex("ItemType") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "ItemType"
            .ColHidden(.ColIndex("ItemType")) = True
        End If
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "UnitID"
        '.ColKey(6) = "UnitID"
  
        .ColPosition(.ColIndex("UnitID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("UnitID")) = "ÊÕœ… «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("UnitID")) = "Item unit"
        End If

        .ColAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
    
        '≈÷«›… «·„ﬁ«”
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemSize"
        .ColPosition(.ColIndex("ItemSize")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemSize")) = "«·„ﬁ«”"
        Else
            .TextMatrix(0, .ColIndex("ItemSize")) = "Item Size"
        End If
    
        .ColAlignment(.ColIndex("ItemSize")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemSize")) = flexAlignRightCenter
        StrSQL = "Select * From TblItemsSizes Order by SizeName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "SizeName", "SizeId")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ItemSize")) = "|" & StrList
        End If
    
        '--------------------------------------------
        '≈ŸÂ«— «··Ê‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ColorID"
        .ColPosition(.ColIndex("ColorID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ColorID")) = "«··Ê‰"
        Else
            .TextMatrix(0, .ColIndex("ColorID")) = "Item Color"
        End If

        .ColAlignment(.ColIndex("ColorID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ColorID")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "ColorName", "ColorID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ColorID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
      
      
         'If SystemOptions.IsGeometricProportions = True Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "UnitPrice"
            .ColPosition(.ColIndex("UnitPrice")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("UnitPrice")) = "”⁄— «·ÊÕœ…"
            Else
                .TextMatrix(0, .ColIndex("UnitPrice")) = "UnitPrice"
            End If
        
            .ColAlignment(.ColIndex("UnitPrice")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("UnitPrice")) = flexAlignRightCenter
        'End If

        '--------------------------------------------
    
        '--------------------------------------------
        '≈ŸÂ«— «·›∆…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ClassId"
        .ColPosition(.ColIndex("ClassId")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ClassId")) = "«·›—“"
        Else
            .TextMatrix(0, .ColIndex("ClassId")) = "Grade"
        End If

        .ColAlignment(.ColIndex("ClassId")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ClassId")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblItemsclasses Order by SizeName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "SizeName", "SizeId")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ClassId")) = "|" & StrList
        End If

        '--------------------------------------------
    
        '--------------------------------------------
        '≈ŸÂ«— „Êﬁ⁄ «· Œ“Ì‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "StoreIDLoc"
        .ColPosition(.ColIndex("StoreIDLoc")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("StoreIDLoc")) = " „Êﬁ⁄ «· Œ“Ì‰"
        Else
            .TextMatrix(0, .ColIndex("StoreIDLoc")) = "Store Locations"
        End If

        .ColAlignment(.ColIndex("StoreIDLoc")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("StoreIDLoc")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblstoresLocations Order by name"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "name", "Locid")

        If StrList <> "" Then
            .ColComboList(.ColIndex("StoreIDLoc")) = "|" & StrList
        End If

        '--------------------------------------------
'****************************************************
         '--------------------------------------------
'****************************************************
      ''////////////////////////////// «·„Œ“‰ ⁄·Ï „” ÊÏ «·Ã—Ìœ
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "StoreID2"
        .ColPosition(.ColIndex("StoreID2")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("StoreID2")) = "«·„Œ“‰"
        Else
            .TextMatrix(0, .ColIndex("StoreID2")) = "Store"
        End If

        .ColAlignment(.ColIndex("StoreID2")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("StoreID2")) = flexAlignRightCenter
        If checkmanyStores(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "Select StoreID,StoreName From TblStore  "
     Else
           StrSQL = "Select StoreID,StoreNamee From TblStore  "
    End If
    End If
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If SystemOptions.UserInterface = ArabicInterface Then

        StrList = .BuildComboList(RsNote, "StoreName", "StoreID")
     Else
         StrList = .BuildComboList(RsNote, "StoreNamee", "StoreID")
   End If

        If StrList <> "" Then
            .ColComboList(.ColIndex("StoreID2")) = "|" & StrList
        End If

        RsNote.Close



'**********************************************************
        '≈ŸÂ«—   «·«Ã—«¡«  «· ’ÕÌÕÌ…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CorrectionID"
        .ColPosition(.ColIndex("CorrectionID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CorrectionID")) = "«·«Ã—«¡«  «· ’ÕÌÕÌ…"
        Else
            .TextMatrix(0, .ColIndex("CorrectionID")) = "Corrective action"
        End If

        .ColAlignment(.ColIndex("CorrectionID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CorrectionID")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblCorrectaction Order by name"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "name", "Cid")

        If StrList <> "" Then
            .ColComboList(.ColIndex("CorrectionID")) = "|" & StrList
        End If

        '--------------------------------------------
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ProductionDate"
        .ColComboList(.ColIndex("ProductionDate")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ProductionDate")) = " «—ÌŒ «·«‰ «Ã"
        Else
            .TextMatrix(0, .ColIndex("ProductionDate")) = "Production Date"
        End If

        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "GranteeData"
        .ColComboList(.ColIndex("GranteeData")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("GranteeData")) = " »Ì«‰«  «·÷„«‰"
        Else
            .TextMatrix(0, .ColIndex("GranteeData")) = "guarantee Data "
        End If
    
    
            .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "GoldData"
        .ColComboList(.ColIndex("GoldData")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("GoldData")) = "  ›«’Ì· «·ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("GoldData")) = "Details "
        End If
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ExpiryDate"
        .ColComboList(.ColIndex("ExpiryDate")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ExpiryDate")) = " «—ÌŒ «·«‰ Â«¡"
        Else
            .TextMatrix(0, .ColIndex("ExpiryDate")) = "Expiry Date"
        End If
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "LotNO"

        If Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Then
            .ColComboList(.ColIndex("LotNO")) = "..."
        End If

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("LotNO")) = "—ﬁ„ «··Êÿ"
        Else
            .TextMatrix(0, .ColIndex("LotNO")) = "Lot NO"
        End If
 
 
 
    
    
 '--------------------------------------------
        '≈ŸÂ«— «·„Ê—œÌ‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "SupplierID"
        .ColPosition(.ColIndex("SupplierID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("SupplierID")) = "«·„Ê—œ"
        Else
            .TextMatrix(0, .ColIndex("SupplierID")) = "Supplier"
        End If

        .ColAlignment(.ColIndex("SupplierID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("SupplierID")) = flexAlignRightCenter
    
           StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

          
                StrSQL = StrSQL + " And (CusID <> 1)"
      
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CusName", "CusID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("SupplierID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
  '*******************************************************
          '--------------------------------------------
        '≈÷«›… ⁄„Êœ   «·ﬂ «·ÊÃ ≈› —«÷Ì« ›Ï ﬂ· «·‘«‘« 
   
        
         '?U?C? C??«·ﬂ «·ÊÃ???
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "catalog"
        .ColPosition(.ColIndex("catalog")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("catalog")) = "«·ﬂ «·ÊÃ"
        Else
            .TextMatrix(0, .ColIndex("catalog")) = "catalog"
        End If

        .ColAlignment(.ColIndex("catalog")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("catalog")) = flexAlignRightCenter
    
            StrSQL = "select * from TblItemCatalog"
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CatlogName", "ID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("catalog")) = "|" & StrList
        End If

        RsNote.Close
        
   '*******************************************************
       
               '≈ŸÂ«— «· ’‰Ì⁄
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "countris"
        .ColPosition(.ColIndex("countris")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("countris")) = "»·œ «·’‰⁄"
        Else
            .TextMatrix(0, .ColIndex("countris")) = "made in"
        End If

        .ColAlignment(.ColIndex("countris")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("countris")) = flexAlignRightCenter
    
            StrSQL = "select * from TblCountriesData"
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CountryName", "CountryID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("countris")) = "|" & StrList
        End If

        RsNote.Close
        
        

            
       
 '--------------------------------------------
        '≈ŸÂ«— «·⁄„·« 
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ScurrencyID"
        .ColPosition(.ColIndex("ScurrencyID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ScurrencyID")) = "«·⁄„·…"
        Else
            .TextMatrix(0, .ColIndex("ScurrencyID")) = "Currency"
        End If

        .ColAlignment(.ColIndex("ScurrencyID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ScurrencyID")) = flexAlignRightCenter
    If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "Select id,name From currency  order by id"
Else
  StrSQL = "Select id,namee From currency  order by id"
End If
          
             
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If SystemOptions.UserInterface = ArabicInterface Then

        StrList = .BuildComboList(RsNote, "name", "id")
Else
StrList = .BuildComboList(RsNote, "namee", "id")
End If

        If StrList <> "" Then
            .ColComboList(.ColIndex("ScurrencyID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
       
       
       
               '≈ŸÂ«—   «·ﬂÊœ «· Õ·Ì·Ì ··ﬁÿ⁄Â
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ParrtNoCode"
        .ColPosition(.ColIndex("ParrtNoCode")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ParrtNoCode")) = "ﬂÊœ «·ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("ParrtNoCode")) = "Part No"
        End If

        .ColAlignment(.ColIndex("ParrtNoCode")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ParrtNoCode")) = flexAlignRightCenter
        
        
                       '≈ŸÂ«—   «·ﬂÊœ «· Õ·Ì·Ì ··ﬁÿ⁄Â
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemDetailedCode"
        .ColPosition(.ColIndex("ItemDetailedCode")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemDetailedCode")) = "ﬂÊœ   Õ·Ì·Ì ··ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("ItemDetailedCode")) = "Part No"
        End If

        .ColAlignment(.ColIndex("ItemDetailedCode")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemDetailedCode")) = flexAlignRightCenter
        
 '*****************
 If Me.GridTrans = purchaseorderrequest Or Me.GridTrans = purchaseorder Or Me.GridTrans = purchaseOrderApproved Then
       If SystemOptions.UserInterface = ArabicInterface Then
                 .ColComboList(.ColIndex("Shipping")) = "#1;  «·„‘ —Ì|#2; «·»«∆⁄"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("Shipping")) = "#1;Buyer |#2;Seller "
            End If
End If
'*********************

 
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Remarks"
        .ColComboList(.ColIndex("Remarks")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Remarks")) = "„·«ÕŸ« "
        Else
            .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        End If
        
        
         .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "QtyFaqtors"
   
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("QtyFaqtors")) = " «·⁄·«ﬁ… „⁄ «·ÊÕœ…"
        Else
            .TextMatrix(0, .ColIndex("QtyFaqtors")) = "Qty Factor"
        End If
         .ColAlignment(.ColIndex("QtyFaqtors")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("QtyFaqtors")) = flexAlignRightCenter
    
'****************************************
     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CostPriceTk"
        .ColPosition(.ColIndex("CostPriceTk")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CostPriceTk")) = "”⁄— «· ﬂ·›… ··«” ·«„"
        Else
            .TextMatrix(0, .ColIndex("CostPriceTk")) = "Cost Price"
        End If

        .ColAlignment(.ColIndex("CostPriceTk")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CostPriceTk")) = flexAlignRightCenter
        
'**************************************************************
'If SystemOptions.AllowWorkWithArea = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist) Then
                               
                               '≈ŸÂ«—     «·«— ›«⁄
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Height"
        .ColPosition(.ColIndex("Height")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Height")) = "«·«— ›«⁄"
        Else
            .TextMatrix(0, .ColIndex("Height")) = "Height"
        End If

        .ColAlignment(.ColIndex("Height")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Height")) = flexAlignRightCenter
        
    '*******************************************************************
                              '≈ŸÂ«—     ﬁÿ— Œ«—ÃÌ
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "OUTR"
        .ColPosition(.ColIndex("OUTR")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("OUTR")) = "ﬁÿ— Œ«—ÃÌ"
        Else
            .TextMatrix(0, .ColIndex("OUTR")) = "OUTR"
        End If

        .ColAlignment(.ColIndex("OUTR")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("OUTR")) = flexAlignRightCenter
        
        
        
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ForUnit"
        .ColPosition(.ColIndex("ForUnit")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ForUnit")) = "·ﬂ· ÊÕœ…"
        Else
            .TextMatrix(0, .ColIndex("ForUnit")) = "OUTR"
        End If

        .ColAlignment(.ColIndex("ForUnit")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ForUnit")) = flexAlignRightCenter
        
        .ColHidden(.ColIndex("ForUnit")) = True
        
        
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "MethodCalc"
        .ColPosition(.ColIndex("MethodCalc")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("MethodCalc")) = "ÿ—Ìﬁ… Õ”«»"
        Else
            .TextMatrix(0, .ColIndex("MethodCalc")) = "MethodCalc"
        End If

        .ColAlignment(.ColIndex("MethodCalc")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("MethodCalc")) = flexAlignRightCenter
        
        .ColHidden(.ColIndex("MethodCalc")) = True
        
        
               
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "lowering"
        .ColPosition(.ColIndex("lowering")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("lowering")) = " ‰ﬁÌ’"
        Else
            .TextMatrix(0, .ColIndex("lowering")) = "MethodCalc"
        End If

        .ColAlignment(.ColIndex("lowering")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("lowering")) = flexAlignRightCenter
        
        .ColHidden(.ColIndex("lowering")) = True
        
               
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "increase"
        .ColPosition(.ColIndex("increase")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("increase")) = "“Ì«œ…"
        Else
            .TextMatrix(0, .ColIndex("increase")) = "MethodCalc"
        End If

        .ColAlignment(.ColIndex("increase")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("increase")) = flexAlignRightCenter
        
        .ColHidden(.ColIndex("increase")) = True
        
      '*******************************************************************
        
       '*******************************************************************
                              '≈ŸÂ«—     ﬁÿ— œ«Œ·Ì
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "INR"
        .ColPosition(.ColIndex("INR")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("INR")) = "ﬁÿ— œ«Œ·Ì"
        Else
            .TextMatrix(0, .ColIndex("INR")) = "OUTR"
        End If

        .ColAlignment(.ColIndex("INR")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("INR")) = flexAlignRightCenter
        
      '*******************************************************************
          
      
        
      '*******************************************************************
          
          
        
      '*******************************************************************
                              '≈ŸÂ«—     «·ÿÊ·
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "length"
        .ColPosition(.ColIndex("length")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("length")) = "«·ÿÊ·"
        Else
            .TextMatrix(0, .ColIndex("length")) = "length"
        End If

        .ColAlignment(.ColIndex("length")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("length")) = flexAlignRightCenter
        
      '*******************************************************************
        
              '*******************************************************************
                              '≈ŸÂ«—     «·ÿÊ·
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemQty"
        .ColPosition(.ColIndex("ItemQty")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemQty")) = "«·ÿÊ·"
        Else
            .TextMatrix(0, .ColIndex("ItemQty")) = "length"
        End If

        .ColAlignment(.ColIndex("ItemQty")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemQty")) = flexAlignRightCenter
        .ColHidden(.ColIndex("ItemQty")) = True
        
                                       '≈ŸÂ«—     «·⁄—÷
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Width"
        .ColPosition(.ColIndex("Width")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Width")) = "«·⁄—÷"
        Else
            .TextMatrix(0, .ColIndex("Width")) = "Width"
        End If

        .ColAlignment(.ColIndex("Width")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Width")) = flexAlignRightCenter
        
                                               '≈ŸÂ«—     «·„”«Õ…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Area"
        .ColPosition(.ColIndex("Area")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Area")) = "«·„”«Õ…"
        Else
            .TextMatrix(0, .ColIndex("Area")) = "Area"
        End If

        .ColAlignment(.ColIndex("Area")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Area")) = flexAlignRightCenter
                                                       '≈ŸÂ«—     «·⁄œœ
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "NoCount"
        .ColPosition(.ColIndex("NoCount")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("NoCount")) = "«·⁄œœ"
        Else
            .TextMatrix(0, .ColIndex("NoCount")) = "No"
        End If

        .ColAlignment(.ColIndex("NoCount")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("NoCount")) = flexAlignRightCenter
'End If







        
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "LastPurchasePrice"
        .ColPosition(.ColIndex("LastPurchasePrice")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("LastPurchasePrice")) = "«Œ— ”⁄— ‘—«¡"
        Else
            .TextMatrix(0, .ColIndex("LastPurchasePrice")) = "Last Purchase Price"
        End If

        .ColAlignment(.ColIndex("LastPurchasePrice")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("LastPurchasePrice")) = flexAlignRightCenter
        
                        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemBalance"
        .ColPosition(.ColIndex("ItemBalance")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemBalance")) = "—’Ìœ «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("ItemBalance")) = "Item Balance"
        End If

        .ColAlignment(.ColIndex("ItemBalance")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemBalance")) = flexAlignRightCenter
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemSalesYear"
        .ColPosition(.ColIndex("ItemSalesYear")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemSalesYear")) = "ÕÃ„ „»Ì⁄«  «·’‰› Œ·«· ⁄«„"
        Else
            .TextMatrix(0, .ColIndex("ItemSalesYear")) = "Item Sales Year"
        End If

        .ColAlignment(.ColIndex("ItemSalesYear")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemSalesYear")) = flexAlignRightCenter
        
     
        ''//////////////////ÃœÌœ «Ê „ œ«Ê·
          .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "New_Rolling"
        .ColPosition(.ColIndex("New_Rolling")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("New_Rolling")) = "Õ«·… «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("New_Rolling")) = "Item Status"
        End If
          If SystemOptions.UserInterface = ArabicInterface Then
                 .ColComboList(.ColIndex("New_Rolling")) = "#1;  ÃœÌœ|#2; „ œ«Ê·"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("New_Rolling")) = "#1;New |#2;Rolling "
            End If
        .ColAlignment(.ColIndex("New_Rolling")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("New_Rolling")) = flexAlignRightCenter

         '≈ŸÂ«— „Êﬁ⁄ Õœ «·ÿ·»
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ReminRequ"
        .ColPosition(.ColIndex("ReminRequ")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ReminRequ")) = "»«ﬁÌ Õœ «·ÿ·»"
        Else
            .TextMatrix(0, .ColIndex("ReminRequ")) = "Remaining Request"
        End If

        .ColAlignment(.ColIndex("ReminRequ")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ReminRequ")) = flexAlignRightCenter
                '*******************************************19032017*********
                             '******************************
              .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "IsExpirDate"
        .ColPosition(.ColIndex("IsExpirDate")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("IsExpirDate")) = "   «—ÌŒ ÌœÊÌ"
        Else
            .TextMatrix(0, .ColIndex("IsExpirDate")) = "Date Manual"
        End If

        .ColAlignment(.ColIndex("IsExpirDate")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("IsExpirDate")) = flexAlignRightCenter
        
             '******************************
             
                        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "L"
        .ColPosition(.ColIndex("L")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("L")) = "L"
        Else
            .TextMatrix(0, .ColIndex("L")) = "L"
        End If

        .ColAlignment(.ColIndex("L")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("L")) = flexAlignRightCenter
      ''////
           
           
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "W"
        .ColPosition(.ColIndex("W")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("W")) = "W"
        Else
            .TextMatrix(0, .ColIndex("W")) = "W"
        End If

        .ColAlignment(.ColIndex("W")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("W")) = flexAlignRightCenter
   
              '''''///
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "H1"
        .ColPosition(.ColIndex("H1")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("H1")) = "H1"
        Else
            .TextMatrix(0, .ColIndex("H1")) = "H1"
        End If

        .ColAlignment(.ColIndex("H1")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("H1")) = flexAlignRightCenter
        
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "H2"
        .ColPosition(.ColIndex("H2")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("H2")) = "H2"
        Else
            .TextMatrix(0, .ColIndex("H2")) = "H2"
        End If

        .ColAlignment(.ColIndex("H2")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("H2")) = flexAlignRightCenter
        
             '******************************
                         .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Vatyo"
        .ColPosition(.ColIndex("Vatyo")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Vatyo")) = "‰”»… «· VAT %"
        Else
            .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        End If

        .ColAlignment(.ColIndex("Vatyo")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Vatyo")) = flexAlignRightCenter


                     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Vat"
        .ColPosition(.ColIndex("Vat")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Vat")) = "ﬁÌ„… «·  VAT"
        Else
            .TextMatrix(0, .ColIndex("Vat")) = "VAT "
        End If

        .ColAlignment(.ColIndex("Vat")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Vat")) = flexAlignRightCenter
        


                     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "TypeVAT"
        .ColPosition(.ColIndex("TypeVAT")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("TypeVAT")) = "‰Ê⁄ «· VAT "
        Else
            .TextMatrix(0, .ColIndex("TypeVAT")) = "VAT Type"
        End If

        .ColAlignment(.ColIndex("TypeVAT")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("TypeVAT")) = flexAlignRightCenter
        
        
     '********************************************
     
             .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "MixNo"
        .ColPosition(.ColIndex("MixNo")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("MixNo")) = "ﬂÊœ «·„ﬂ”"
        Else
            .TextMatrix(0, .ColIndex("MixNo")) = "Mix. No"
        End If

        .ColAlignment(.ColIndex("MixNo")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("MixNo")) = flexAlignRightCenter


'***********************************************************************
         .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "MaxQty"
        .ColPosition(.ColIndex("MaxQty")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("MaxQty")) = " ﬂ„Ì… «·ﬂ»—Ï"
        Else
            .TextMatrix(0, .ColIndex("MaxQty")) = "Max Qty"
        End If

        .ColAlignment(.ColIndex("MaxQty")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("MaxQty")) = flexAlignRightCenter
    ''//////////////
            '≈ŸÂ«— «ﬂ»— ÊÕœ…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "MaxUnitID"
        .ColPosition(.ColIndex("MaxUnitID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("MaxUnitID")) = "«·ÊÕœ… «·ﬂ»—Ï"
        Else
            .TextMatrix(0, .ColIndex("MaxUnitID")) = "Max Unit"
        End If

        .ColAlignment(.ColIndex("MaxUnitID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("MaxUnitID")) = flexAlignRightCenter
    
           StrSQL = "Select * From TblUnites "

        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
        StrList = .BuildComboList(RsNote, "UnitName", "UnitID")
        Else
        StrList = .BuildComboList(RsNote, "UnitNamee", "UnitID")
        End If
        If StrList <> "" Then
            .ColComboList(.ColIndex("MaxUnitID")) = "|" & StrList
        End If

        RsNote.Close
        



              
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CusID2"
        .ColPosition(.ColIndex("CusID2")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CusID2")) = "«”„ «·⁄„Ì·"
        Else
            .TextMatrix(0, .ColIndex("CusID2")) = "Customer Name"
        End If

        .ColAlignment(.ColIndex("CusID2")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CusID2")) = flexAlignRightCenter
    
           StrSQL = " SELECT     CusID, CusName, CusNamee"
           StrSQL = StrSQL & "            From dbo.TblCustemers"
           StrSQL = StrSQL & "           WHERE     (Type = 1) "
           If Not isCashCust Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If
            StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"

        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
        StrList = .BuildComboList(RsNote, "CusName", "CusID")
        Else
        StrList = .BuildComboList(RsNote, "CusNamee", "CusID")
        End If
        If StrList <> "" Then
            .ColComboList(.ColIndex("CusID2")) = "|" & StrList
        End If

        RsNote.Close




     '**********************************************************************
              .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "NoHours"
        .ColPosition(.ColIndex("NoHours")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("NoHours")) = "⁄œœ «·”«⁄« "
        Else
            .TextMatrix(0, .ColIndex("NoHours")) = "No.Hours"
        End If

        .ColAlignment(.ColIndex("NoHours")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("NoHours")) = flexAlignRightCenter
        ''''
                  .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "PriceNoHours"
        .ColPosition(.ColIndex("PriceNoHours")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("PriceNoHours")) = " ﬂ·›… «·”«⁄…"
        Else
            .TextMatrix(0, .ColIndex("PriceNoHours")) = "Cost Hour"
        End If

        .ColAlignment(.ColIndex("PriceNoHours")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("PriceNoHours")) = flexAlignRightCenter
        ''''
        
                      .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "TotalPriceNoHours"
        .ColPosition(.ColIndex("TotalPriceNoHours")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("TotalPriceNoHours")) = "«Ã„«·Ì  ﬂ·›… «·”«⁄« "
        Else
            .TextMatrix(0, .ColIndex("TotalPriceNoHours")) = "Total Cost Hours"
        End If

        .ColAlignment(.ColIndex("TotalPriceNoHours")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("TotalPriceNoHours")) = flexAlignRightCenter
        
        
        
        
        
                              .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "TotalInvoiceQty"
        .ColPosition(.ColIndex("TotalInvoiceQty")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("TotalInvoiceQty")) = "«·ﬂ„Ì… «·«Ã„«·Ì…"
        Else
            .TextMatrix(0, .ColIndex("TotalInvoiceQty")) = "Total Qty"
        End If

        .ColAlignment(.ColIndex("TotalInvoiceQty")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("TotalInvoiceQty")) = flexAlignRightCenter
        
        
                              .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ISSUEDQTY"
        .ColPosition(.ColIndex("ISSUEDQTY")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ISSUEDQTY")) = "«·„‰’—› ”«»ﬁ«"
        Else
            .TextMatrix(0, .ColIndex("ISSUEDQTY")) = "ISSUED Qty"
        End If

        .ColAlignment(.ColIndex("ISSUEDQTY")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ISSUEDQTY")) = flexAlignRightCenter
                
                
        
        '**********************************************************************
                ''//////////
              ''////
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ManufactCompany"
        .ColPosition(.ColIndex("ManufactCompany")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ManufactCompany")) = "«·‘—ﬂ… «·’«‰⁄…"
        Else
            .TextMatrix(0, .ColIndex("ManufactCompany")) = "Manufacturing Company"
        End If

        .ColAlignment(.ColIndex("ManufactCompany")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ManufactCompany")) = flexAlignRightCenter
              '''''///
                      ''////
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CountryOrg"
        .ColPosition(.ColIndex("CountryOrg")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CountryOrg")) = " »·œ «·„‰‘√"
        Else
            .TextMatrix(0, .ColIndex("CountryOrg")) = "Country of Origin"
        End If

        .ColAlignment(.ColIndex("CountryOrg")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CountryOrg")) = flexAlignRightCenter
              '''''///

        '***********************************************************************
                   '≈ŸÂ«—  «·„ÊŸ›
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "EmpID4"
        .ColPosition(.ColIndex("EmpID4")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("EmpID4")) = " «”„ «·„ÊŸ›"
        Else
            .TextMatrix(0, .ColIndex("EmpID4")) = "Employee Name "
        End If

        .ColAlignment(.ColIndex("EmpID4")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("EmpID4")) = flexAlignRightCenter
    
'           StrSQL = "Select * From TblEmployee "

   If SystemOptions.UserInterface = ArabicInterface Then
      StrSQL = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee"
   Else
   StrSQL = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name"
   End If
    StrSQL = StrSQL & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
       
    StrSQL = StrSQL & "     where  ( JobTypeName like '%⁄«„·%'  or JobTypeNamee like '%Labor%' )  "
 


        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If SystemOptions.UserInterface = ArabicInterface Then
        StrList = .BuildComboList(RsNote, "Emp_Name", "Emp_ID")
        Else
        StrList = .BuildComboList(RsNote, "Emp_Namee", "Emp_ID")
        End If
        If StrList <> "" Then
            .ColComboList(.ColIndex("EmpID4")) = "|" & StrList
        End If

        RsNote.Close
 
        '***********************************************************************
        
             .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "SalesValue"
        .ColPosition(.ColIndex("SalesValue")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("SalesValue")) = "”⁄— »Ì⁄ «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("SalesValue")) = "Sales Value"
        End If

        .ColAlignment(.ColIndex("SalesValue")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("SalesValue")) = flexAlignRightCenter



     ''///////////////
   .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "TotalPrice"
        .ColPosition(.ColIndex("TotalPrice")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("TotalPrice")) = "«·”⁄— «·‘«„·"
        Else
            .TextMatrix(0, .ColIndex("TotalPrice")) = "Total"
        End If

        .ColAlignment(.ColIndex("TotalPrice")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("TotalPrice")) = flexAlignRightCenter
        ''''''
        
        

Dim MyStrList As String
        Dim LngItemID As Long

        '     StrSQL = "SELECT * FROM TblUnites "
        '                Set Rs = New ADODB.Recordset
        '                Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        '                If Not (Rs.BOF Or Rs.EOF) Then
        '                    MyStrList = .BuildComboList(Rs, "UnitName", "UnitID")
        '                    Grid.ComboList = MyStrList
        '               End If

        'Set the Grid Drag Drop Cap
        '   MsgBox .ColIndex("UnitID")
        '-------------------------------------------
        'Set the Grid tool bar
        SetToolBar

        '-------------------------------------------
        If Not Me.LblItemsCount Is Nothing Then
            If SystemOptions.UserInterface = ArabicInterface Then
                LblItemsCount.ToolTipText = "⁄œœ «·√’‰«› «·„ÊÃÊœ… ›Ï «·›« Ê—…."
            Else
                LblItemsCount.ToolTipText = "Count of items in the invoice."
            End If
        End If

    End With

    '--------------------------------------
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
'        StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set ItemType=0 Where ItemType IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set RelatedItem=0 Where RelatedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
xxx:

    Set RsItems = New ADODB.Recordset
    StrSQL = "Select * From TblItems  where 1=1"

StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.TblItems.HaveSerial, dbo.TblItems.LastUpdate, "
StrSQL = StrSQL & "                       dbo.TblItems.UserID, dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblItems.RequestLimit, dbo.TblItems.CustomerPrice, dbo.TblItems.DealerPrice,"
StrSQL = StrSQL & "                       dbo.TblItems.HaveGuarantee, dbo.TblItems.GuaranteeValue, dbo.TblItems.GuaranteeType, dbo.TblItems.IsArchive, dbo.TblItems.ItemType, dbo.TblItems.AssbliedItem,"
StrSQL = StrSQL & "                       dbo.TblItems.RelatedItem, dbo.TblItems.ItemComment, dbo.TblItems.ItemCase, dbo.TblItems.ItemMaking, dbo.TblItems.ItemMakingNew, dbo.TblItems.code,"
StrSQL = StrSQL & "                       dbo.TblItems.Branch_NO, dbo.TblItems.Fullcode, dbo.TblItems.prifix, dbo.TblItems.PartNo, dbo.TblItems.CostPrice, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                       dbo.TblItems.DefaultSupplier, dbo.TblItems.itemSerials, dbo.TblItems.BinLocation, dbo.TblItems.minvalueqty, dbo.TblItems.MaxValueqty, dbo.TblItems.FreeQty,"
StrSQL = StrSQL & "                       dbo.TblItems.barCodeNO, dbo.TblItems.CatlogNO, dbo.TblItems.FactoryNO, dbo.TblItems.TemplateID, dbo.TblItems.ItemMaxDiscount, dbo.TblItems.OverHead,"
StrSQL = StrSQL & "                       dbo.TblItems.Wight , dbo.TblItems.Content, dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.maxRecivePeriod"
StrSQL = StrSQL & "  FROM         dbo.TblItems  where 1=1 "
  If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowOnlyItemsOfSales = True Then
    StrSQL = StrSQL & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    StrSQL = StrSQL & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
  End If
  
       Dim ContID As Double
       
        
    If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowItemByCustomer = True Then
        If Not Customer Is Nothing Then
            ContID = CheckCustomerCont(val(Customer.BoundText), DtpBillDate.value)
        End If
   
        If ContID <> 0 Then
            StrSQL = StrSQL & "  and dbo.TblItems.ItemID in(   SELECT     ItemID"
            StrSQL = StrSQL & " From dbo.TblCustomerContractDetails"
            StrSQL = StrSQL & " Where (TblCustomerContractD = " & ContID & "))"
        End If
    End If
'****************************************
If SystemOptions.WorkWithLINKEDiTEMS = True Then

        If Me.TxtModFlag = "N" Or Me.TxtModFlag = "E" Then
       
        
                If val(StoreName.BoundText) <> -1 Then
    
'02 06 2015
            StrSQL = GETLINKSQL(val(StoreName.BoundText), 0, Me.TxtModFlag)

                GoTo order_By
            
                End If
        
        End If
        
        
        
        
End If


'**********************************


    If Me.GridTrans = Destruction Then
        StrSQL = StrSQL
        GoTo order_By
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ShowPrice Then
        StrSQL = StrSQL + " and (TblItems.AssbliedItem=1 OR TblItems.AssbliedItem=0)"
    ElseIf Me.GridTrans = ProductionOrder Then
        StrSQL = StrSQL + " and (TblItems.ItemMakingNew=1)"
    Else
        StrSQL = StrSQL + " and (TblItems.AssbliedItem=0)"
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        StrSQL = StrSQL + " AND (TblItems.ItemType=0 OR TblItems.ItemType=1)"
    Else
        StrSQL = StrSQL + " AND (TblItems.ItemType=0)"
    End If
'StrSQL = StrSQL + " AND (ItemID<=3500)"

order_By:

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order By ItemName"
    Else
        StrSQL = StrSQL + " Order By ItemNamee"
    End If

    RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsItems.BOF Or RsItems.EOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrList = Me.Grid.BuildComboList(RsItems, "ItemName", "ItemID")
        Else
            StrList = Me.Grid.BuildComboList(RsItems, "ItemNamee", "ItemID")
        End If

        If StrList <> "" Then
            Grid.ColComboList(Grid.ColIndex("Name")) = "|" & StrList
        End If

         StrList = Grid.BuildComboList(RsItems, "barCodeNO", "ItemID")

        If StrList <> "" Then
            Grid.ColComboList(Grid.ColIndex("Code")) = "|" & StrList
        End If

        RsItems.MoveFirst
Else
Grid.ColComboList(Grid.ColIndex("Name")) = "|" & "0"
Grid.ColComboList(Grid.ColIndex("Code")) = "|" & "0"

    End If

    '--------------------------------------
    FillItemData (storename1)
    Dim mFirstTabIndex As Integer
    mFirstTabIndex = Me.Grid.TabIndex
    If Not TxtCustCode Is Nothing Then
        TxtCustCode.TabIndex = mFirstTabIndex + 1
        
        Customer.TabIndex = TxtCustCode.TabIndex + 1
        mFirstTabIndex = Customer.TabIndex
    End If
    DCboItemCode.TabIndex = mFirstTabIndex + 1
    DCboItemName.TabIndex = Me.DCboItemCode.TabIndex + 1
    CboItemCase.TabIndex = DCboItemName.TabIndex + 1
    
    If Not TxtSerial Is Nothing Then
        TxtSerial.TabIndex = CboItemCase.TabIndex + 1
        TxtQuantity.TabIndex = TxtSerial.TabIndex + 1
    Else
        TxtQuantity.TabIndex = CboItemCase.TabIndex + 1
    End If

    TxtPrice.TabIndex = TxtQuantity.TabIndex + 1
    CmdAddData.TabIndex = TxtPrice.TabIndex + 1

    '---------------------------------
    If Not Me.LblItemsCount Is Nothing Then
        LblItemsCount.backcolor = &H404040
        LblItemsCount.ForeColor = &HFFFF&
        LblItemsCount.Alignment = vbCenter
        LblItemsCount.RightToLeft = True
    End If
If storename1 = -1 Then
    GetMySetting
End If
    Exit Sub
ErrTrap:
    'Stop
    'Resume
End Sub


Public Sub FillGridx22112017(Optional storename1 As Integer = -1)
firstrun = True
    Dim StrList As String
    Dim RsNote As ADODB.Recordset
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
   ' On Error GoTo ErrTrap

    m_ShowQtyToolTip = True
If storename1 <> -1 Then GoTo xxx
    With Grid

        If SystemOptions.UserInterface = EnglishInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
            .TextMatrix(0, .ColIndex("Code")) = "Item Code"
            .TextMatrix(0, .ColIndex("Name")) = "Item Name"

      If .ColIndex("NProductionOrderNO") <> -1 Then
                .TextMatrix(0, .ColIndex("NProductionOrderNO")) = "Pro. Order No"
            End If
            
            
            If .ColIndex("Serial") <> -1 Then
                .TextMatrix(0, .ColIndex("Serial")) = "Item Serial"
            End If

            If .ColIndex("ItemCase") <> -1 Then
                .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case"
            Else
        
            End If

            If .ColIndex("HaveSerial") <> -1 Then
                .TextMatrix(0, .ColIndex("HaveSerial")) = "Have SerialS"
            Else
                .TextMatrix(0, .ColIndex("HaveSerial")) = "Have SerialS"

            End If
        
            '.TextMatrix(0, .ColIndex("UnitID")) = "Unit"
         
            .TextMatrix(0, .ColIndex("Count")) = "Quantity"
            .TextMatrix(0, .ColIndex("Price")) = "Price"

            If .ColIndex("Valu") <> -1 Then
                .TextMatrix(0, .ColIndex("Valu")) = "Item Value"
            End If

            If .ColIndex("DiscountVal") <> -1 Then
                .TextMatrix(0, .ColIndex("DiscountVal")) = "Discount Value"
            End If

            If .ColIndex("DiscountType") <> -1 Then
                .TextMatrix(0, .ColIndex("DiscountType")) = "Discount Type"
            End If

            If .ColIndex("guaranteeTime") <> -1 Then
                .TextMatrix(0, .ColIndex("guaranteeTime")) = "Guarantee Interval"
            End If
        
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
            .ColComboList(.ColIndex("ItemCase")) = "#1;ÃœÌœ|#2;„” ⁄„·"
        Else
            .ColComboList(.ColIndex("ItemCase")) = "#1;New|#2;Used"
        End If


        '--------------------------------------------------------------------------
        If m_GridTrans <> OpeningBalance And m_GridTrans <> NewGard Then
            If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("DiscountType")) = "#1;·«ÌÊÃœ Œ’„|#2;Œ’„ »ﬁÌ„…|#3;Œ’„ »‰”»…|#4;„Ã«‰Ì"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("DiscountType")) = "#1;NO Discount|#2;Value Discount|#3;Percent Discount|#4;Free"
            End If
           If .ColIndex("ExtraType") <> -1 Then
        
                  If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("ExtraType")) = "#1;·«ÌÊÃœ ⁄„Ê·…|#2; ⁄„Ê·… ﬁÌ„…|#3;⁄„Ê·…  ‰”»… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("ExtraType")) = "#1;NA|#2;Value commision|#3;Percent commision "
            End If
           End If
            
        
        End If
    
        .ColAlignment(.ColIndex("Serial")) = flexAlignLeftCenter
        '--------------------------------------------
        .Cols = .Cols + 1
        .ColWidthMin = 350
        .ColKey(.Cols - 1) = "AddSerials"
        .ColComboList(.ColIndex("AddSerials")) = "..."
    
        .ColPosition(.ColIndex("AddSerials")) = 5 ' .ColPosition(.ColIndex("Serail")/
    
        '--------------------------------------------
        '≈÷«›… ⁄„Êœ ‰Ê⁄ «·’‰› ≈› —«÷Ì« ›Ï ﬂ· «·‘«‘« 
        If .ColIndex("ItemType") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "ItemType"
            .ColHidden(.ColIndex("ItemType")) = True
        End If
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "UnitID"
        '.ColKey(6) = "UnitID"
  
        .ColPosition(.ColIndex("UnitID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("UnitID")) = "ÊÕœ… «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("UnitID")) = "Item unit"
        End If

        .ColAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
    
        '≈÷«›… «·„ﬁ«”
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemSize"
        .ColPosition(.ColIndex("ItemSize")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemSize")) = "«·„ﬁ«”"
        Else
            .TextMatrix(0, .ColIndex("ItemSize")) = "Item Size"
        End If
    
        .ColAlignment(.ColIndex("ItemSize")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemSize")) = flexAlignRightCenter
        StrSQL = "Select * From TblItemsSizes Order by SizeName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "SizeName", "SizeId")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ItemSize")) = "|" & StrList
        End If
    
        '--------------------------------------------
        '≈ŸÂ«— «··Ê‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ColorID"
        .ColPosition(.ColIndex("ColorID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ColorID")) = "«··Ê‰"
        Else
            .TextMatrix(0, .ColIndex("ColorID")) = "Item Color"
        End If

        .ColAlignment(.ColIndex("ColorID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ColorID")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "ColorName", "ColorID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ColorID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
      
        '--------------------------------------------
    
        '--------------------------------------------
        '≈ŸÂ«— «·›∆…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ClassId"
        .ColPosition(.ColIndex("ClassId")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ClassId")) = "«·›—“"
        Else
            .TextMatrix(0, .ColIndex("ClassId")) = "Grade"
        End If

        .ColAlignment(.ColIndex("ClassId")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ClassId")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblItemsclasses Order by SizeName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "SizeName", "SizeId")

        If StrList <> "" Then
            .ColComboList(.ColIndex("ClassId")) = "|" & StrList
        End If

        '--------------------------------------------
    
        '--------------------------------------------
        '≈ŸÂ«— „Êﬁ⁄ «· Œ“Ì‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "StoreIDLoc"
        .ColPosition(.ColIndex("StoreIDLoc")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("StoreIDLoc")) = " „Êﬁ⁄ «· Œ“Ì‰"
        Else
            .TextMatrix(0, .ColIndex("StoreIDLoc")) = "Store Locations"
        End If

        .ColAlignment(.ColIndex("StoreIDLoc")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("StoreIDLoc")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblstoresLocations Order by name"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "name", "Locid")

        If StrList <> "" Then
            .ColComboList(.ColIndex("StoreIDLoc")) = "|" & StrList
        End If

        '--------------------------------------------
'****************************************************
      ''////////////////////////////// «·„Œ“‰ ⁄·Ï „” ÊÏ «·Ã—Ìœ
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "StoreID2"
        .ColPosition(.ColIndex("StoreID2")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("StoreID2")) = "«·„Œ“‰"
        Else
            .TextMatrix(0, .ColIndex("StoreID2")) = "Store"
        End If

        .ColAlignment(.ColIndex("StoreID2")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("StoreID2")) = flexAlignRightCenter
    If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "Select StoreID,StoreName From TblStore  "
     Else
           StrSQL = "Select StoreID,StoreNamee From TblStore  "
    End If
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If SystemOptions.UserInterface = ArabicInterface Then

        StrList = .BuildComboList(RsNote, "StoreName", "StoreID")
     Else
         StrList = .BuildComboList(RsNote, "StoreNamee", "StoreID")
   End If

        If StrList <> "" Then
            .ColComboList(.ColIndex("StoreID2")) = "|" & StrList
        End If

        RsNote.Close


'**********************************************************
        '≈ŸÂ«—   «·«Ã—«¡«  «· ’ÕÌÕÌ…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CorrectionID"
        .ColPosition(.ColIndex("CorrectionID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CorrectionID")) = "«·«Ã—«¡«  «· ’ÕÌÕÌ…"
        Else
            .TextMatrix(0, .ColIndex("CorrectionID")) = "Corrective action"
        End If

        .ColAlignment(.ColIndex("CorrectionID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CorrectionID")) = flexAlignRightCenter
    
        StrSQL = "Select * From TblCorrectaction Order by name"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "name", "Cid")

        If StrList <> "" Then
            .ColComboList(.ColIndex("CorrectionID")) = "|" & StrList
        End If

        '--------------------------------------------
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ProductionDate"
        .ColComboList(.ColIndex("ProductionDate")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ProductionDate")) = " «—ÌŒ «·«‰ «Ã"
        Else
            .TextMatrix(0, .ColIndex("ProductionDate")) = "Production Date"
        End If

        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "GranteeData"
        .ColComboList(.ColIndex("GranteeData")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("GranteeData")) = " »Ì«‰«  «·÷„«‰"
        Else
            .TextMatrix(0, .ColIndex("GranteeData")) = "guarantee Data "
        End If
    
    
            .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "GoldData"
        .ColComboList(.ColIndex("GoldData")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("GoldData")) = "  ›«’Ì· «·ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("GoldData")) = "Details "
        End If
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ExpiryDate"
        .ColComboList(.ColIndex("ExpiryDate")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ExpiryDate")) = " «—ÌŒ «·«‰ Â«¡"
        Else
            .TextMatrix(0, .ColIndex("ExpiryDate")) = "Expiry Date"
        End If
    
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "LotNO"

        If Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Then
            .ColComboList(.ColIndex("LotNO")) = "..."
        End If

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("LotNO")) = "—ﬁ„ «··Êÿ"
        Else
            .TextMatrix(0, .ColIndex("LotNO")) = "Lot NO"
        End If
 
 
 
    
    
 '--------------------------------------------
        '≈ŸÂ«— «·„Ê—œÌ‰
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "SupplierID"
        .ColPosition(.ColIndex("SupplierID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("SupplierID")) = "«·„Ê—œ"
        Else
            .TextMatrix(0, .ColIndex("SupplierID")) = "Supplier"
        End If

        .ColAlignment(.ColIndex("SupplierID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("SupplierID")) = flexAlignRightCenter
    
           StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

          
                StrSQL = StrSQL + " And (CusID <> 1)"
      
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CusName", "CusID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("SupplierID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
  '*******************************************************
          '--------------------------------------------
        '≈÷«›… ⁄„Êœ   «·ﬂ «·ÊÃ ≈› —«÷Ì« ›Ï ﬂ· «·‘«‘« 
   
        
         '?U?C? C??«·ﬂ «·ÊÃ???
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "catalog"
        .ColPosition(.ColIndex("catalog")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("catalog")) = "«·ﬂ «·ÊÃ"
        Else
            .TextMatrix(0, .ColIndex("catalog")) = "catalog"
        End If

        .ColAlignment(.ColIndex("catalog")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("catalog")) = flexAlignRightCenter
    
            StrSQL = "select * from TblItemCatalog"
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CatlogName", "ID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("catalog")) = "|" & StrList
        End If

        RsNote.Close
        
   '*******************************************************
       
               '≈ŸÂ«— «· ’‰Ì⁄
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "countris"
        .ColPosition(.ColIndex("countris")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("countris")) = "»·œ «·’‰⁄"
        Else
            .TextMatrix(0, .ColIndex("countris")) = "made in"
        End If

        .ColAlignment(.ColIndex("countris")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("countris")) = flexAlignRightCenter
    
            StrSQL = "select * from TblCountriesData"
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        StrList = .BuildComboList(RsNote, "CountryName", "CountryID")

        If StrList <> "" Then
            .ColComboList(.ColIndex("countris")) = "|" & StrList
        End If

        RsNote.Close
        
        

            
       
 '--------------------------------------------
        '≈ŸÂ«— «·⁄„·« 
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ScurrencyID"
        .ColPosition(.ColIndex("ScurrencyID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ScurrencyID")) = "«·⁄„·…"
        Else
            .TextMatrix(0, .ColIndex("ScurrencyID")) = "Currency"
        End If

        .ColAlignment(.ColIndex("ScurrencyID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ScurrencyID")) = flexAlignRightCenter
    If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "Select id,name From currency  order by id"
Else
  StrSQL = "Select id,namee From currency  order by id"
End If
          
             
            
    '    StrSQL = "Select * From TblItemsColors Order by ColorName"
        Set RsNote = New ADODB.Recordset
        RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If SystemOptions.UserInterface = ArabicInterface Then

        StrList = .BuildComboList(RsNote, "name", "id")
Else
StrList = .BuildComboList(RsNote, "namee", "id")
End If

        If StrList <> "" Then
            .ColComboList(.ColIndex("ScurrencyID")) = "|" & StrList
        End If

        RsNote.Close
        '--------------------------------------------
       
       
       
               '≈ŸÂ«—   «·ﬂÊœ «· Õ·Ì·Ì ··ﬁÿ⁄Â
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ParrtNoCode"
        .ColPosition(.ColIndex("ParrtNoCode")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ParrtNoCode")) = "ﬂÊœ «·ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("ParrtNoCode")) = "Part No"
        End If

        .ColAlignment(.ColIndex("ParrtNoCode")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ParrtNoCode")) = flexAlignRightCenter
        
        
                       '≈ŸÂ«—   «·ﬂÊœ «· Õ·Ì·Ì ··ﬁÿ⁄Â
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemDetailedCode"
        .ColPosition(.ColIndex("ItemDetailedCode")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemDetailedCode")) = "ﬂÊœ   Õ·Ì·Ì ··ﬁÿ⁄Â"
        Else
            .TextMatrix(0, .ColIndex("ItemDetailedCode")) = "Part No"
        End If

        .ColAlignment(.ColIndex("ItemDetailedCode")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemDetailedCode")) = flexAlignRightCenter
        
 '*****************
 If Me.GridTrans = purchaseorderrequest Or Me.GridTrans = purchaseorder Or Me.GridTrans = purchaseOrderApproved Then
       If SystemOptions.UserInterface = ArabicInterface Then
                 .ColComboList(.ColIndex("Shipping")) = "#1;  «·„‘ —Ì|#2; «·»«∆⁄"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("Shipping")) = "#1;Buyer |#2;Seller "
            End If
End If
'*********************

 
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Remarks"
        .ColComboList(.ColIndex("Remarks")) = "..."

        '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Remarks")) = "„·«ÕŸ« "
        Else
            .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        End If
        
        
 
    
    
'****************************************
     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "CostPriceTk"
        .ColPosition(.ColIndex("CostPriceTk")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("CostPriceTk")) = "”⁄— «· ﬂ·›… ··«” ·«„"
        Else
            .TextMatrix(0, .ColIndex("CostPriceTk")) = "Cost Price"
        End If

        .ColAlignment(.ColIndex("CostPriceTk")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("CostPriceTk")) = flexAlignRightCenter
        
'**************************************************************
'If SystemOptions.AllowWorkWithArea = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist) Then
                               
                               '≈ŸÂ«—     «·«— ›«⁄
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Height"
        .ColPosition(.ColIndex("Height")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Height")) = "«·«— ›«⁄"
        Else
            .TextMatrix(0, .ColIndex("Height")) = "Height"
        End If

        .ColAlignment(.ColIndex("Height")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Height")) = flexAlignRightCenter
        
                                       '≈ŸÂ«—     «·⁄—÷
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Width"
        .ColPosition(.ColIndex("Width")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Width")) = "«·⁄—÷"
        Else
            .TextMatrix(0, .ColIndex("Width")) = "Width"
        End If

        .ColAlignment(.ColIndex("Width")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Width")) = flexAlignRightCenter
        
                                               '≈ŸÂ«—     «·„”«Õ…
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Area"
        .ColPosition(.ColIndex("Area")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Area")) = "«·„”«Õ…"
        Else
            .TextMatrix(0, .ColIndex("Area")) = "Area"
        End If

        .ColAlignment(.ColIndex("Area")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Area")) = flexAlignRightCenter
                                                       '≈ŸÂ«—     «·⁄œœ
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "NoCount"
        .ColPosition(.ColIndex("NoCount")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("NoCount")) = "«·⁄œœ"
        Else
            .TextMatrix(0, .ColIndex("NoCount")) = "No"
        End If

        .ColAlignment(.ColIndex("NoCount")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("NoCount")) = flexAlignRightCenter
'End If







        
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "LastPurchasePrice"
        .ColPosition(.ColIndex("LastPurchasePrice")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("LastPurchasePrice")) = "«Œ— ”⁄— ‘—«¡"
        Else
            .TextMatrix(0, .ColIndex("LastPurchasePrice")) = "Last Purchase Price"
        End If

        .ColAlignment(.ColIndex("LastPurchasePrice")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("LastPurchasePrice")) = flexAlignRightCenter
        
                        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemBalance"
        .ColPosition(.ColIndex("ItemBalance")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemBalance")) = "—’Ìœ «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("ItemBalance")) = "Item Balance"
        End If

        .ColAlignment(.ColIndex("ItemBalance")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemBalance")) = flexAlignRightCenter
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemSalesYear"
        .ColPosition(.ColIndex("ItemSalesYear")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ItemSalesYear")) = "ÕÃ„ „»Ì⁄«  «·’‰› Œ·«· ⁄«„"
        Else
            .TextMatrix(0, .ColIndex("ItemSalesYear")) = "Item Sales Year"
        End If

        .ColAlignment(.ColIndex("ItemSalesYear")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ItemSalesYear")) = flexAlignRightCenter
        
     
        ''//////////////////ÃœÌœ «Ê „ œ«Ê·
          .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "New_Rolling"
        .ColPosition(.ColIndex("New_Rolling")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("New_Rolling")) = "Õ«·… «·’‰›"
        Else
            .TextMatrix(0, .ColIndex("New_Rolling")) = "Item Status"
        End If
          If SystemOptions.UserInterface = ArabicInterface Then
                 .ColComboList(.ColIndex("New_Rolling")) = "#1;  ÃœÌœ|#2; „ œ«Ê·"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("New_Rolling")) = "#1;New |#2;Rolling "
            End If
        .ColAlignment(.ColIndex("New_Rolling")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("New_Rolling")) = flexAlignRightCenter

         '≈ŸÂ«— „Êﬁ⁄ Õœ «·ÿ·»
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ReminRequ"
        .ColPosition(.ColIndex("ReminRequ")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("ReminRequ")) = "»«ﬁÌ Õœ «·ÿ·»"
        Else
            .TextMatrix(0, .ColIndex("ReminRequ")) = "Remaining Request"
        End If

        .ColAlignment(.ColIndex("ReminRequ")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("ReminRequ")) = flexAlignRightCenter
                '*******************************************19032017*********
                             '******************************
              .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "IsExpirDate"
        .ColPosition(.ColIndex("IsExpirDate")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("IsExpirDate")) = "   «—ÌŒ ÌœÊÌ"
        Else
            .TextMatrix(0, .ColIndex("IsExpirDate")) = "Date Manual"
        End If

        .ColAlignment(.ColIndex("IsExpirDate")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("IsExpirDate")) = flexAlignRightCenter
        
             '******************************
             
                        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "L"
        .ColPosition(.ColIndex("L")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("L")) = "L"
        Else
            .TextMatrix(0, .ColIndex("L")) = "L"
        End If

        .ColAlignment(.ColIndex("L")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("L")) = flexAlignRightCenter
      ''////
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "W"
        .ColPosition(.ColIndex("W")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("W")) = "W"
        Else
            .TextMatrix(0, .ColIndex("W")) = "W"
        End If

        .ColAlignment(.ColIndex("W")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("W")) = flexAlignRightCenter
        
              '''''///
              
              ''////

        
              '''''///
              
           .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "H1"
        .ColPosition(.ColIndex("H1")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("H1")) = "H1"
        Else
            .TextMatrix(0, .ColIndex("H1")) = "H1"
        End If

        .ColAlignment(.ColIndex("H1")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("H1")) = flexAlignRightCenter
        
        
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "H2"
        .ColPosition(.ColIndex("H2")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("H2")) = "H2"
        Else
            .TextMatrix(0, .ColIndex("H2")) = "H2"
        End If

        .ColAlignment(.ColIndex("H2")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("H2")) = flexAlignRightCenter
        
             '******************************
             

                     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Vat"
        .ColPosition(.ColIndex("Vat")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Vat")) = "«·ﬁÌ„… «·„÷«›…"
        Else
            .TextMatrix(0, .ColIndex("Vat")) = "Vat "
        End If

        .ColAlignment(.ColIndex("Vat")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Vat")) = flexAlignRightCenter
        


                     .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "Vatyo"
        .ColPosition(.ColIndex("Vatyo")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Vatyo")) = "«·ﬁÌ„… «·„÷«›… %"
        Else
            .TextMatrix(0, .ColIndex("Vatyo")) = "Vat …%"
        End If

        .ColAlignment(.ColIndex("Vatyo")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("Vatyo")) = flexAlignRightCenter
        
        
     '********************************************
Dim MyStrList As String
        Dim LngItemID As Long

        '     StrSQL = "SELECT * FROM TblUnites "
        '                Set Rs = New ADODB.Recordset
        '                Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        '                If Not (Rs.BOF Or Rs.EOF) Then
        '                    MyStrList = .BuildComboList(Rs, "UnitName", "UnitID")
        '                    Grid.ComboList = MyStrList
        '               End If

        'Set the Grid Drag Drop Cap
        '   MsgBox .ColIndex("UnitID")
        '-------------------------------------------
        'Set the Grid tool bar
        SetToolBar

        '-------------------------------------------
        If Not Me.LblItemsCount Is Nothing Then
            If SystemOptions.UserInterface = ArabicInterface Then
                LblItemsCount.ToolTipText = "⁄œœ «·√’‰«› «·„ÊÃÊœ… ›Ï «·›« Ê—…."
            Else
                LblItemsCount.ToolTipText = "Count of items in the invoice."
            End If
        End If

    End With

    '--------------------------------------
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
'        StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set ItemType=0 Where ItemType IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
'        StrSQL = "Update TblItems Set RelatedItem=0 Where RelatedItem IS NULL"
'        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
xxx:

    Set RsItems = New ADODB.Recordset
    StrSQL = "Select * From TblItems  where 1=1"

StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.TblItems.HaveSerial, dbo.TblItems.LastUpdate, "
StrSQL = StrSQL & "                       dbo.TblItems.UserID, dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblItems.RequestLimit, dbo.TblItems.CustomerPrice, dbo.TblItems.DealerPrice,"
StrSQL = StrSQL & "                       dbo.TblItems.HaveGuarantee, dbo.TblItems.GuaranteeValue, dbo.TblItems.GuaranteeType, dbo.TblItems.IsArchive, dbo.TblItems.ItemType, dbo.TblItems.AssbliedItem,"
StrSQL = StrSQL & "                       dbo.TblItems.RelatedItem, dbo.TblItems.ItemComment, dbo.TblItems.ItemCase, dbo.TblItems.ItemMaking, dbo.TblItems.ItemMakingNew, dbo.TblItems.code,"
StrSQL = StrSQL & "                       dbo.TblItems.Branch_NO, dbo.TblItems.Fullcode, dbo.TblItems.prifix, dbo.TblItems.PartNo, dbo.TblItems.CostPrice, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                       dbo.TblItems.DefaultSupplier, dbo.TblItems.itemSerials, dbo.TblItems.BinLocation, dbo.TblItems.minvalueqty, dbo.TblItems.MaxValueqty, dbo.TblItems.FreeQty,"
StrSQL = StrSQL & "                       dbo.TblItems.barCodeNO, dbo.TblItems.CatlogNO, dbo.TblItems.FactoryNO, dbo.TblItems.TemplateID, dbo.TblItems.ItemMaxDiscount, dbo.TblItems.OverHead,"
StrSQL = StrSQL & "                       dbo.TblItems.Wight , dbo.TblItems.Content, dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.maxRecivePeriod"
StrSQL = StrSQL & "  FROM         dbo.TblItems  where 1=1 "
                       
                      
'****************************************
If SystemOptions.WorkWithLINKEDiTEMS = True Then

        If Me.TxtModFlag = "N" Or Me.TxtModFlag = "E" Then
                If val(StoreName.BoundText) <> -1 Then
    
'02 06 2015
            StrSQL = GETLINKSQL(val(StoreName.BoundText), 0, Me.TxtModFlag)

                GoTo order_By
            
                End If
        
        End If
        
        
        
        
End If


'**********************************


    If Me.GridTrans = Destruction Then
        StrSQL = StrSQL
        GoTo order_By
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ShowPrice Then
        StrSQL = StrSQL + " and (TblItems.AssbliedItem=1 OR TblItems.AssbliedItem=0)"
    ElseIf Me.GridTrans = ProductionOrder Then
        StrSQL = StrSQL + " and (TblItems.ItemMakingNew=1)"
    Else
        StrSQL = StrSQL + " and (TblItems.AssbliedItem=0)"
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        StrSQL = StrSQL + " AND (TblItems.ItemType=0 OR TblItems.ItemType=1)"
    Else
        StrSQL = StrSQL + " AND (TblItems.ItemType=0)"
    End If
'StrSQL = StrSQL + " AND (ItemID<=3500)"

order_By:

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order By ItemName"
    Else
        StrSQL = StrSQL + " Order By ItemNamee"
    End If

    RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsItems.BOF Or RsItems.EOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrList = Me.Grid.BuildComboList(RsItems, "ItemName", "ItemID")
        Else
            StrList = Me.Grid.BuildComboList(RsItems, "ItemNamee", "ItemID")
        End If

        If StrList <> "" Then
            Grid.ColComboList(Grid.ColIndex("Name")) = "|" & StrList
        End If

         StrList = Grid.BuildComboList(RsItems, "barCodeNO", "ItemID")

        If StrList <> "" Then
            Grid.ColComboList(Grid.ColIndex("Code")) = "|" & StrList
        End If

        RsItems.MoveFirst
Else
Grid.ColComboList(Grid.ColIndex("Name")) = "|" & "0"
Grid.ColComboList(Grid.ColIndex("Code")) = "|" & "0"

    End If

    '--------------------------------------
    FillItemData (storename1)
    DCboItemCode.TabIndex = Me.Grid.TabIndex + 1
    DCboItemName.TabIndex = Me.DCboItemCode.TabIndex + 1
    CboItemCase.TabIndex = DCboItemName.TabIndex + 1

    If Not TxtSerial Is Nothing Then
        TxtSerial.TabIndex = CboItemCase.TabIndex + 1
        TxtQuantity.TabIndex = TxtSerial.TabIndex + 1
    Else
        TxtQuantity.TabIndex = CboItemCase.TabIndex + 1
    End If

    TxtPrice.TabIndex = TxtQuantity.TabIndex + 1
    CmdAddData.TabIndex = TxtPrice.TabIndex + 1

    '---------------------------------
    If Not Me.LblItemsCount Is Nothing Then
        LblItemsCount.backcolor = &H404040
        LblItemsCount.ForeColor = &HFFFF&
        LblItemsCount.Alignment = vbCenter
        LblItemsCount.RightToLeft = True
    End If
If storename1 = -1 Then
    GetMySetting
End If
    Exit Sub
ErrTrap:
    'Stop
    'Resume
End Sub

Private Sub Grid_BeforeSort(ByVal Col As Long, _
                            Order As Integer)
    Dim i As Integer

    With Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) = "" Then
            .RemoveItem .Rows - 1
            Calculate 1
        End If

    End With

End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, _
                                 ByVal Col As Long)
    Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim rdate As Date
    Dim Frm As FrmGridAddItemComment
    'Dim Frm1 As ItemProductionDate

    On Error GoTo ErrTrap

    With Me.Grid

        Select Case .ColKey(Col)

            Case "NewItem"
                On Error GoTo ErrTrap

                If Me.TxtModFlag.Text = "N" Or Me.TxtModFlag.Text = "E" Then
                 '   FrmAddNewItem.DealingForm = Me.GridTrans
       '  FrmAddNewItem.show vbModal
                    Unload FrmItems
                    FrmItems.DealingForm = Me.GridTrans
                    FrmItems.CALLEDFPRM = True
                    FrmItems.show
              FrmItems.Cmd_Click (0)
                  '  FrmItems.ZOrder 0
                FrmItems.SetFocus
                
                    
                End If

            Case "AddSerials"

                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Then
                    If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexUnchecked Then
                        Exit Sub
                    End If

                    LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
                    LngStoreID = val(val(Me.StoreName.BoundText))

                    If LngItemID = 0 Or LngStoreID = 0 Then
                        Exit Sub
                    End If

                    PopUpSerialList LngItemID, Row, Me.Grid.ColIndex("Serial")
                End If

            Case "Remarks"
                ItemComment Row, Col
         
            Case "ProductionDate"
                ItemProductionDate Row, Col, True, 0
            
            Case "GranteeData"
                ItemGranteeData Row, Col, True
                  Case "GoldData"
                ItemGoldData Row, Col, True
            Case "ExpiryDate"
                ItemProductionDate Row, Col, , 1
            
            Case "LotNO"
           '     LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
           '     LngStoreID = val(val(Me.storename.BoundText))
           '     rdate = DtpBillDate.Value
           '     ItemLotNoFun Row, Col, LngItemID, LngStoreID, rdate
            
            
                            LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
                If ChicIsLotNo(LngItemID) = True Then
                LngStoreID = val(Me.StoreName.BoundText)
                rdate = DtpBillDate.value
                'ItemLotNoFun Row, Col, LngItemID, LngStoreID, rdate
                ItemLotNoFun Row, Col, LngItemID, val(.TextMatrix(Row, .ColIndex("StoreID2"))), rdate
                
                
                End If
                
        End Select
        If m_GridTrans = InvoiceTransaction Then
                    If Not CBoBasedON Is Nothing Then
                            If val(CBoBasedON.ListIndex) = 4 Then
                            Exit Sub
                            End If
                    End If
        End If


    End With

ErrTrap:
End Sub

Private Sub Grid_Click()
    'On Error GoTo ErrTrap
    'If Me.GridTrans = InvoiceTransaction Then
    '    With Grid '
    '        If .Col = .ColIndex("Code") Then
    '            If .Cell(flexcpData, .Row, .ColIndex("Replace")) <> "" Then
    '                FrmSaleBill.ReplacementData
    '            End If
    '        End If
    '    End With
    'End If
    'Exit Sub
    'ErrTrap:
End Sub
Public Sub FillItemData(Optional storename1 As Integer = -1)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    Dim IntItemType As Integer
    Dim IntAssbliedType As Integer

    Set Dcombos = New ClsDataCombos

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)

    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›

    'If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = ShowPrice Then
    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
           Dim ContID As Double
           Dim StrItem As String
    If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowItemByCustomer = True Then
   ContID = CheckCustomerCont(val(Customer.BoundText), DtpBillDate.value)
   If ContID <> 0 Then
    StrItem = StrItem & "  and dbo.TblItems.ItemID in(   SELECT     ItemID"
    StrItem = StrItem & " From dbo.TblCustomerContractDetails"
    StrItem = StrItem & " Where (TblCustomerContractD = " & ContID & "))"
  End If
  End If
  If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowOnlyItemsOfSales = True Then
    StrItem = StrItem & "  and dbo.TblItems.GroupID in(   SELECT     GroupID"
    StrItem = StrItem & " From dbo.Groups  WHERE     (ISNULL(POSGroup, 0) = 1))"
  End If
    StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
        Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, , storename1, , , StrItem
        Dcombos.GetItemsNames Me.DCboItemName, -1, -1, , , storename1, , , StrItem
    ElseIf Me.GridTrans = PurchaseTransaction Then

        Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0, , storename1
        Dcombos.GetItemsNames Me.DCboItemName, 0, 0, , , storename1
    ElseIf Me.GridTrans = ProductionOrder Then
        Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, 1, storename1
        Dcombos.GetItemsNames Me.DCboItemName, -1, -1, 1, , storename1

    Else
        Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0, , storename1
        Dcombos.GetItemsNames Me.DCboItemName, 0, 0, , , storename1
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        StrSQL = StrSQL + " AND (TblItems.ItemType=0 OR TblItems.ItemType=1)"
    Else
        StrSQL = StrSQL + " AND TblItems.ItemType=0"
    End If

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DCboItemName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboItemCode

    'Õ«·… «·’‰›
    With CboItemCase
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÃœÌœ"
            .AddItem "„” ⁄„·"
        Else
            .AddItem "New"
            .AddItem "Used"
        End If

    End With

    '‰Ê⁄ «·Œ’„
    'With CboDiscountType
    ''    .AddItem "·«ÌÊÃœ Œ’„"
    ''    .AddItem "Œ’„ »ﬁÌ„…"
    ''    .AddItem "Œ’„ »‰”»…"
    ''    .AddItem "„Ã«‰Ì"
    'End With
    Exit Sub
ErrTrap:
End Sub

Public Sub FillItemDatax22112017(Optional storename1 As Integer = -1)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    Dim IntItemType As Integer
    Dim IntAssbliedType As Integer

    Set Dcombos = New ClsDataCombos

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)

    ' ⁄»∆… »Ì«‰«  «·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
    'ﬂÊœ «·’‰›

    'If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = ShowPrice Then
    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, , storename1
        Dcombos.GetItemsNames Me.DCboItemName, -1, -1, , , storename1
    ElseIf Me.GridTrans = PurchaseTransaction Then

        Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0, , storename1
        Dcombos.GetItemsNames Me.DCboItemName, 0, 0, , , storename1
    ElseIf Me.GridTrans = ProductionOrder Then
        Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, 1, storename1
        Dcombos.GetItemsNames Me.DCboItemName, -1, -1, 1, , storename1

    Else
        Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0, , storename1
        Dcombos.GetItemsNames Me.DCboItemName, 0, 0, , , storename1
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        StrSQL = StrSQL + " AND TblItems.ItemType=0 OR TblItems.ItemType=1"
    Else
        StrSQL = StrSQL + " AND TblItems.ItemType=0"
    End If

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DCboItemName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboItemCode

    'Õ«·… «·’‰›
    With CboItemCase
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "ÃœÌœ"
            .AddItem "„” ⁄„·"
        Else
            .AddItem "New"
            .AddItem "Used"
        End If

    End With

    '‰Ê⁄ «·Œ’„
    'With CboDiscountType
    ''    .AddItem "·«ÌÊÃœ Œ’„"
    ''    .AddItem "Œ’„ »ﬁÌ„…"
    ''    .AddItem "Œ’„ »‰”»…"
    ''    .AddItem "„Ã«‰Ì"
    'End With
    Exit Sub
ErrTrap:
End Sub

Private Sub Grid_DblClick()
    Dim itemcode As String
    Dim ItemName As String

    With Me.Grid

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        
        If .Col = .ColIndex("Serial") And Me.TxtModFlag.Text <> "R" Then
            ItemSSErials Grid.Row, Grid.Col
        End If
  
        If .Col = .ColIndex("Count") Then
      '      If val(.TextMatrix(.Row, .ColIndex("FoxyNo"))) = 0 Then
      '          .TextMatrix(.Row, .ColIndex("FoxyNo")) = setfoxyNo("id2")
      '      End If
    If m_GridTrans = InvoiceTransaction Then
                    If Not CBoBasedON Is Nothing Then
                            If val(CBoBasedON.ListIndex) = 4 Then
                            Exit Sub
                            End If
                    End If
        End If
            
       ItemDetailsNewIgia .Row, .Col
       Exit Sub
            FrmItemsDetails.show
            FrmItemsDetails.TxtItemID = .TextMatrix(.Row, .ColIndex("Code"))
            GetItemData val(.TextMatrix(.Row, .ColIndex("Code"))), itemcode, ItemName
            FrmItemsDetails.txtItemCode = itemcode
            FrmItemsDetails.TxtItemName = ItemName
          
            FrmItemsDetails.txtQty = .TextMatrix(.Row, .ColIndex("Count"))
            FrmItemsDetails.TxtUnitID = .Cell(flexcpData, .Row, .ColIndex("UnitID"))
            FrmItemsDetails.TxtUnitName = .TextMatrix(.Row, .ColIndex("UnitId"))
            

 
            
            FrmItemsDetails.TxtFoxyNo = .TextMatrix(.Row, .ColIndex("FoxyNo"))
            FrmItemsDetails.Retrive FrmItemsDetails.TxtFoxyNo
             
            If Me.TxtModFlag = "R" Then
                FrmItemsDetails.Grid.Editable = flexEDNone
                FrmItemsDetails.Ele(1).Enabled = False
            End If

            '  If Me.GridTrans = PurchaseTransaction Then
            '  FrmItemsDetails.TxtArrivalDate.text = FrmBillBuy.DTArrivalDate.value
            '  End If
        
        End If

    End With

    If Me.TxtModFlag.Text = "R" Then

        With Me.Grid

            If .Row = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            If .Col = .ColIndex("Seria1l") Then
                If Trim(.TextMatrix(.Row, .ColIndex("Serial"))) = "" Then
                    Exit Sub
                End If

                Load FrmSerialData

                With FrmSerialData
                    .XPTxtCode.Text = Trim(Grid.TextMatrix(Grid.Row, Grid.ColIndex("Serial")))
                    .DCboItemName.BoundText = Me.Grid.TextMatrix(Grid.Row, Grid.ColIndex("Name"))
                    .show
                    .Cmd_Click 0
                    .ZOrder 0
                End With

            ElseIf .Col = .ColIndex("Code") Or .Col = .ColIndex("Name") Or .Col = .ColIndex("ItemCase") Then

                If val(.TextMatrix(.Row, .ColIndex("Name"))) <> 0 Then
                    Load FrmSelectData
                    FrmSelectData.DCboItemName.BoundText = val(.TextMatrix(.Row, .ColIndex("Name")))
                    FrmSelectData.txtItemCode.Text = val(.TextMatrix(.Row, .ColIndex("Name")))

                    If Not Me.StoreName Is Nothing Then
                        FrmSelectData.DcboStores.BoundText = val(Me.StoreName.BoundText)
                    End If

                    FrmSelectData.show
                End If

            ElseIf .Col = .ColIndex("Remarks") And Me.TxtModFlag.Text = "R" Then
                ItemComment Grid.Row, Grid.Col

            ElseIf .Col = .ColIndex("ProductionDate") And Me.TxtModFlag.Text = "R" Then
                ItemProductionDate Grid.Row, Grid.Col
            ElseIf .Col = .ColIndex("Serial") And Me.TxtModFlag.Text = "R" Then
            
                ItemSSErials Grid.Row, Grid.Col

            ElseIf .Col = .ColIndex("GranteeData") And Me.TxtModFlag.Text = "R" Then
        
                ItemGranteeData Grid.Row, Grid.Col
    ElseIf .Col = .ColIndex("GoldData") And Me.TxtModFlag.Text = "R" Then
        
                ItemGoldData Grid.Row, Grid.Col
                
            End If

        End With

    End If

End Sub

 

Private Sub Grid_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
    On Error Resume Next
    Dim currentrow As Integer
    Dim LngItemID As Long
    If KeyCode = vbKeyReturn Then

        With Grid
        
If SystemOptions.workWithBarcode = True And Me.GridTrans = InvoiceTransaction Then

   If val(.TextMatrix(.Row, .ColIndex("Price"))) <> 0 Then
                currentrow = .Row
                '.Rows = .Rows + 1
                .Row = .Row + 1

                If currentrow = .Row Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                End If
         
                .Col = .ColIndex("Code")
          
        
      End If
      
GoTo ll
End If

        
            If .Col = .ColIndex("Price") And val(.TextMatrix(.Row, .ColIndex("Price"))) <> 0 Then
                currentrow = .Row
                '.Rows = .Rows + 1
                .Row = .Row + 1

                If currentrow = .Row Then
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                End If
         
                .Col = .ColIndex("Code")
          
            ElseIf .Col = .ColIndex("Code") Then
                .Col = .ColIndex("Name")
            ElseIf .Col = .ColIndex("Name") Then
                .Col = .ColIndex("Unitid")
            ElseIf .Col = .ColIndex("Unitid") Then
                .Col = .ColIndex("Count")
            ElseIf .Col = .ColIndex("Count") Then
                .Col = .ColIndex("Price")
            
   
    
            End If
ll:
            .ShowCell .Row, .Col
            
            .SetFocus
        End With

    End If



   If KeyCode = vbKeyF3 Then

        With Grid
        If .ColKey(.Col) <> "MixNo" Then
        
                                If .Col = .ColIndex("NProductionOrderNO") Then
                                    .Col = .ColIndex("NProductionOrderNO")
                                  LngItemID = val(.TextMatrix(.Row, .ColIndex("Code")))
                                    
                              ProductionSearchFun .Row, .Col, 0
                              
                           Else
                            
                                ItemSearchFun .Row, .Col, 0
                    
                        End If
            End If
        End With
        

    End If

End Sub

Private Sub Grid_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    Dim LngMouseRow As Long
    Dim StrItemSerial As String
    Dim LngStoreID As Long

    If Button = vbRightButton Then

        With Me.Grid
            LngMouseRow = .MouseRow

            If LngMouseRow = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            If val(.TextMatrix(LngMouseRow, .ColIndex("Name"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("Name")))

                If Me.Grid.Cell(flexcpChecked, LngMouseRow, .ColIndex("HaveSerial")) = flexChecked Then
                    If .TextMatrix(LngMouseRow, .ColIndex("Serial")) <> "..." Or .TextMatrix(LngMouseRow, .ColIndex("Serial")) <> "" Then
                        StrItemSerial = Trim$(.TextMatrix(LngMouseRow, .ColIndex("Serial")))
                    End If

                Else
                    StrItemSerial = ""
                End If

                If Not Me.StoreName Is Nothing Then
                    LngStoreID = val(val(Me.StoreName.BoundText))
                End If

                SetItemReportsMenu LngCurrentItemID, Me.Grid.Parent, StrItemSerial, LngStoreID
            End If

        End With

    End If

End Sub

Private Sub Grid_OLEDragDrop(Data As VSFlex8UCtl.VSDataObject, _
                             effect As Long, _
                             ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)

    Dim StrTempDropID As String
    Dim LngRow As Long
    On Local Error GoTo ErrTrap

    If Button = 0 Then ' vbLeftButton
        StrTempDropID = Data.GetData(vbCFText)

        With Me.Grid

            If StrTempDropID = "" Then
                effect = vbDropEffectNone
            Else
                LngRow = .MouseRow

                If LngRow = -1 Or LngRow = 0 Then 'test if  any row under mouse curser
                    LngRow = NewEmptyRow
                End If

                effect = vbDropEffectNone

                If LngRow = -1 Then Exit Sub
                AddItemInGrid LngRow, val(StrTempDropID)
                'Unload FrmItemTip
            End If

        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim MyStrList As String
    Dim LngItemID As Long

    With Me.Grid
        '    If Me.GridTrans = InvoiceTransaction Then

        Select Case .ColKey(Col)
        Case "Count"
     
If SystemOptions.WorkWithItemsDetails = True And (Me.GridTrans <> NewGard) Then
              Cancel = True
     End If
     
        
        Case "DiscountVal"
        
  '      If mdifrmmain.POSTRansactiosG.Visible = True Then
'frmmangerlogon.show vbModal
'Exit Sub
'End If
            Case "Price"

             '   If SystemOptions.UserInvoiceChangePrice = 0 Then
                                   If Me.GridTrans = InvoiceTransaction And SystemOptions.UserInvoiceChangePrice = 0 Then
                                 
                                      Cancel = True
                              ' ’—›
                           ElseIf Me.GridTrans = InventoryOut And SystemOptions.UserInvoiceChangePrice1 = 0 Then
                                 
                                      Cancel = True
                               
                                
                                ' «” ·«„
                                      ElseIf (Me.GridTrans = 18 Or Me.GridTrans = 1) And SystemOptions.UserInvoiceChangePrice2 = 0 Then
                                 
                                      Cancel = True
                               
              
                
                Else

                    If m_ShowQtyToolTip = True Then
                        If .TextMatrix(Row, .ColIndex("Code")) <> "" Then
                           ShowItemTip ItemPriceToolTip, .TextMatrix(Row, .ColIndex("Code")), Row, Col
                        End If
                    End If
                End If
            
            
            
            Case ""

            Case "UnitID"
                LngItemID = val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))

                If LngItemID = 0 Then
                    Cancel = True
                Else
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName,TblUnites.UnitNamee "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        Else
                            MyStrList = .BuildComboList(rs, "UnitNamee", "UnitID")
                        End If

                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitID")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
            
            
           
                
                
                  Case "catalog"
                LngItemID = val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))

                If LngItemID = 0 Then
                    Cancel = True
                Else
                    StrSQL = "SELECT ID,CatlogName  "
                    StrSQL = StrSQL + " FROM TblItemCatalog  "
                    StrSQL = StrSQL + " Where ItemID=" & LngItemID
                    
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                   
                            MyStrList = .BuildComboList(rs, "CatlogName", "ID")
                       

                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("catalog")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                    End If
                    
        End Select

        '    End If
    End With

End Sub

Private Sub LblDiscountsTotal_MouseMove(Button As Integer, _
                                        Shift As Integer, _
                                        X As Single, _
                                        Y As Single)
    LblDiscountsTotal.ToolTipText = WriteNo(LblDiscountsTotal.Caption, 0, True)
End Sub

Private Sub LblTotalAll_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    LblTotalAll.ToolTipText = WriteNo(LblTotalAll.Caption, 0, True)
End Sub

Private Sub m_FrmItemsNewPrices_Unload(Cancel As Integer)
    Set m_FrmItemsNewPrices = Nothing
End Sub
Private Sub StoreName_Change()
' Exit Sub
    'On Error GoTo ErrTrap
    Dim i As Long
     Dim Dcombos As ClsDataCombos
 
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
              With Grid
              For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("ExpiryDate")) = ""
              If val(.TextMatrix(i, .ColIndex("Code"))) <> 0 Then
              GetItemBalanceInStore i, 1, 1, 1, , val(.TextMatrix(i, .ColIndex("Code"))), DtpBillDate.value
              End If
                Next i
                
               End With
                                  
                       If SystemOptions.WorkWithLINKEDiTEMS = True Then
                    
                          
                              If StoreName.BoundText = "" Then Exit Sub
                                   With Grid
                                   .Rows = 2
                                  .Clear flexClearScrollable, flexClearEverything
                                .Refresh
                                  End With
                                     FillGrid (val(StoreName.BoundText))
                                  
                        
                             End If
          
         
      ElseIf Me.TxtModFlag.Text = "R" Then
      
       If SystemOptions.WorkWithLINKEDiTEMS = True Then
                    
          '    FillGrid (val(StoreName.BoundText))
        End If
      
    End If
    ColorExperDate
    FillStoredLin
   
    SentTypeVAT
    
End Sub
Private Sub StoreName_Changex19062017()
' Exit Sub
    'On Error GoTo ErrTrap
     Dim Dcombos As ClsDataCombos
 
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
                       If SystemOptions.WorkWithLINKEDiTEMS = True Then
                    
                          
                              If StoreName.BoundText = "" Then Exit Sub
                                            
                                           Grid.Rows = 2
                                Grid.Clear flexClearScrollable, flexClearEverything
                                Grid.Refresh
            
                                     FillGrid (val(StoreName.BoundText))
                                  
                        
                             End If
          
         
      ElseIf Me.TxtModFlag.Text = "R" Then
      
       If SystemOptions.WorkWithLINKEDiTEMS = True Then
                    
              FillGrid (val(StoreName.BoundText))
        End If
        
                    
      
    End If
End Sub

Private Sub STORENAME_Click(Area As Integer)
    'On Error GoTo ErrTrap
 Exit Sub
     Dim Dcombos As ClsDataCombos
 
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
         If SystemOptions.WorkWithLINKEDiTEMS = True Then
       '   Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, , val(storename.BoundText)
       ' Dcombos.GetItemsNames Me.DcboItemName, -1, -1, , val(storename.BoundText)
       If StoreName.BoundText = "" Then Exit Sub
                 
      '         Grid.Rows = 2
    'Grid.Clear flexClearScrollable, flexClearEverything
    'Grid.Refresh
    
        FillGrid (val(StoreName.BoundText))
        
         End If

         
    End If
    SentTypeVAT
End Sub

Private Sub TxtAddTax_Change()

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
        Me.Calculate 1, , , True
    End If

End Sub
Sub RelinVatGrid()
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
End If
Next i
End With
TxtValueAdded.Text = SmValu
End Sub
Private Sub TxtAddTax_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAddTax.Text, 0)
End Sub



Private Sub TxtDiscount_Val_Change()
SentTypeVAT
End Sub
Private Sub GetQtyCodeFormBarCode()

    Dim mFirstPart As String
    Dim mMiddlePart As String
    Dim mLastPart As String
    Dim mResult As String
    Dim mRes As Double
        
    Dim mFromNo As Double
    mFromNo = SystemOptions.FromNo
    
    CodeFromBarCode = TxtItemCodeB1
    QtyFromBarCode = 0
    If SystemOptions.IsSomeItemWeight Then
         
        mRes = val(SystemOptions.CodeTo) - val(SystemOptions.CodeFrom)
        mFirstPart = left(Trim(TxtItemCodeB1), Len(CStr(SystemOptions.FromNo)))
        If mFirstPart <> SystemOptions.FromNo Then
            mFirstPart = left(Trim(TxtItemCodeB1), Len(CStr(SystemOptions.OrNo)))
        End If
        If mFirstPart <> SystemOptions.OrNo And mFirstPart <> SystemOptions.FromNo Then
            GoTo ExitSub
        End If
        If Len(Trim(TxtItemCodeB1)) <> 13 Then GoTo ExitSub
        mResult = Replace(Trim(TxtItemCodeB1), mFirstPart, "", 1, 1)
       ' mResult = Replace(TxtItemCodeB1, mFirstPart, "")
        mMiddlePart = left(mResult, mRes + 1)
        mLastPart = Replace(Trim(TxtItemCodeB1), mFirstPart & mMiddlePart, "", 1, 1)
        mLastPart = left(mLastPart, Len(mLastPart) - 1)
    End If
    If mMiddlePart = "" Then GoTo ExitSub
    CodeFromBarCode = mMiddlePart
    QtyFromBarCode = mLastPart
    Exit Sub
ExitSub:
    CodeFromBarCode = Trim(TxtItemCodeB1)
    QtyFromBarCode = 1
End Sub
'Private Sub TxtDiscount_Val_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'SentTypeVAT
'End If
'End Sub

'Private Sub TxtDiscount_Val_LostFocus()
'If val(CboDiscount_Type.ListIndex) > 0 Then
'SentTypeVAT
'End If
'End Sub

Public Sub TxtItemCodeB1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 '   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
Dim StrSerial As String
Dim rsDummy As New ADODB.Recordset
Dim mQty As Double
CodeFromBarCode = TxtItemCodeB1
QtyFromBarCode = 1

     StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1) AND (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & TxtItemCodeB1 & "'   )"
 Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.EOF Then
    GetQtyCodeFormBarCode
    
         StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
    StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
    StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1) AND (dbo.TblItems.Fullcode ='" & CodeFromBarCode & "'  OR"
    StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & CodeFromBarCode & "'  Or dbo.TblItems.barCodeNO ='" & val(CodeFromBarCode) & "' )"
    RsTemp.Close
     Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
End If

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
  '      LngItemID = val(DCboItemCode.BoundText)
        LngItemID = IIf(IsNull(RsTemp("itemid").value), 0, RsTemp("ItemID").value)
  LngUnitID = IIf(IsNull(RsTemp("UnitID").value), 0, RsTemp("UnitID").value)
  ColorID = 1 'IIf(IsNull(RsTemp("ColorID").Value), 0, RsTemp("ColorID").Value)
  sizeid = 1 'IIf(IsNull(RsTemp("SizeID").Value), 0, RsTemp("SizeID").Value)
 ClassId = 1 ' IIf(IsNull(RsTemp("ClassId").Value), 0, RsTemp("ClassId").Value)
 '   ParrtNoCode = IIf(IsNull(RsTemp("ParrtNoCode").Value), "", RsTemp("ParrtNoCode").Value)
 '       ItemDetailedCode = IIf(IsNull(RsTemp("ItemDetailedCode").Value), "", RsTemp("ItemDetailedCode").Value)
        
        If LngItemID <> 0 Then
        DCboItemName.BoundText = LngItemID
          DCboItemCode.BoundText = LngItemID
           Me.CboItemCase.ListIndex = 0
         If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Or Me.GridTrans = InvoiceTransaction Then
             TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
             
             StrSQL = " SELECT tu.UnitName,tu.UnitNamee FROM TblUnites AS tu Where UnitId = " & LngUnitID
             rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Not rsDummy.EOF Then
                If Trim(rsDummy!UnitName & "") = "Õ»…" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PEICE" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PP" Then
                    Dim mQtyString As String
                    mQtyString = left(CStr(QtyFromBarCode), 2)
                    mQty = val(mQtyString)
                Else
                   mQty = QtyFromBarCode / 1000
                End If
             End If
             
           ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                     If Not TxtSerial Is Nothing Then
                    StrSerial = Trim$(TxtSerial.Text)
                    End If
                     
 
            '        unitid = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                     TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                 
                End If
             TxtQuantity.Text = mQty
            TxtQuantity.Tag = "F"
            TxtQuantity.Enabled = True
        '    Me.Grid.Rows = Me.Grid.Rows + 1
        
    With Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
        End If
         
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = ProductionOrder Then
                   AddItemInGrid .Rows - 1, val(DCboItemName.BoundText), val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode
        End If
'FrmCustomerDisplay.LblInformation.Caption = DCboItemName.Text & vbNewLine & "  Price   " & val(Me.TxtPrice.Text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
'   FrmCustomerDisplay.lblitemid = LngItemID
     End With
           
           
 
         'DCboItemCode_KeyDown vbKeyReturn, 0
          Me.TxtItemCodeB1.Text = ""
          'Unload FrmItemSearch2
      Me.TxtItemCodeB1.SetFocus
         
    Else
        
        
        
        
        
        
        
        
        Exit Sub
    End If
 Else
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "·«ÌÊÃœ ’‰› »Â–« «·ﬂÊœ", vbCritical
 Else
 MsgBox "There is no item in this code", vbCritical
 End If
 
 TxtItemCodeB1.Text = ""
 
    
End If

End If
End Sub
 
 Public Sub TxtItemCodeB_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 '   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String


    StrSQL = " SELECT     ItemDetailedCode, ParrtNoCode, ProductionDate, ExpireDate, ColorID, UnitID, SizeID, ClassId, ItemId"
StrSQL = StrSQL & " from dbo.ItemsDetails"
StrSQL = StrSQL & " GROUP BY ItemDetailedCode, ParrtNoCode, ColorID, UnitID, SizeID, ClassId, ProductionDate, ExpireDate, ItemId"
StrSQL = StrSQL & "  HAVING        (ParrtNoCode = '" & TxtItemCodeB & "')"
     
     
       
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
  '      LngItemID = val(DCboItemCode.BoundText)
        LngItemID = IIf(IsNull(RsTemp("itemid").value), 0, RsTemp("ItemID").value)
  LngUnitID = IIf(IsNull(RsTemp("UnitID").value), 0, RsTemp("UnitID").value)
  ColorID = IIf(IsNull(RsTemp("ColorID").value), 0, RsTemp("ColorID").value)
  sizeid = IIf(IsNull(RsTemp("SizeID").value), 0, RsTemp("SizeID").value)
 ClassId = IIf(IsNull(RsTemp("ClassId").value), 0, RsTemp("ClassId").value)
    ParrtNoCode = IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        ItemDetailedCode = IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID <> 0 Then
        DCboItemName.BoundText = LngItemID
          DCboItemCode.BoundText = LngItemID
           Me.CboItemCase.ListIndex = 0
             TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID)
             TxtQuantity.Text = 1
            TxtQuantity.Tag = "F"
            TxtQuantity.Enabled = True
        '    Me.Grid.Rows = Me.Grid.Rows + 1
        
    With Grid

        If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
        End If
                   AddItemInGrid .Rows - 1, val(DCboItemName.BoundText), val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode
FrmCustomerDisplay.LblInformation.Caption = DCboItemName.Text & vbNewLine & "  Price   " & val(Me.TxtPrice.Text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
   FrmCustomerDisplay.lblitemid = LngItemID
     End With
           
           
 
         'DCboItemCode_KeyDown vbKeyReturn, 0
          Me.TxtItemCodeB.Text = ""
          Unload FrmItemSearch2
      Me.TxtItemCodeB.SetFocus
         
    Else
        
        
        
        
        
        
        
        
        Exit Sub
    End If
    
    
End If

End If
End Sub
 Private Sub TxtModFlag_Change()
    Dim i  As Integer

    If Me.TxtModFlag.Text = "R" Then
    If m_GridTrans <> ProductionOrder And m_GridTrans <> ProductionMaterialReciveVoucher Then
        Me.Grid.OLEDropMode = flexOLEDropNone
        Me.Grid.Editable = flexEDNone
        Else
         Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse
        End If

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = False
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = True
            'Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "E" Then
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = False
            '  Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "N" Then

        If Not Me.LblTotalAll Is Nothing Then
            Me.LblTotalAll.Caption = 0
        End If

        If Not Me.txtTotal Is Nothing Then
            Me.txtTotal.Text = 0
        End If

        If Not Me.LblTotalQty Is Nothing Then
            Me.LblTotalQty.Caption = 0
        End If
    
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            '        Me.GrdTBar.Buttons("Barcode").Enabled = False
            'Me.GrdTBar.Buttons("Options").Enabled = True
            '        Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If
    End If
    
End Sub
Private Sub TxtModFlag_Change22042018()
    Dim i  As Integer

    If Me.TxtModFlag.Text = "R" Then
    If m_GridTrans <> ProductionOrder Then
        Me.Grid.OLEDropMode = flexOLEDropNone
        Me.Grid.Editable = flexEDNone
        Else
         Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse
        End If

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = False
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = True
            'Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "E" Then
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = False
            '  Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "N" Then

        If Not Me.LblTotalAll Is Nothing Then
            Me.LblTotalAll.Caption = 0
        End If

        If Not Me.txtTotal Is Nothing Then
            Me.txtTotal.Text = 0
        End If

        If Not Me.LblTotalQty Is Nothing Then
            Me.LblTotalQty.Caption = 0
        End If
    
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            '        Me.GrdTBar.Buttons("Barcode").Enabled = False
            'Me.GrdTBar.Buttons("Options").Enabled = True
            '        Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If
    End If
    
End Sub
Private Sub TxtModFlag_Changex22012018()
    Dim i  As Integer

    If Me.TxtModFlag.Text = "R" Then
        Me.Grid.OLEDropMode = flexOLEDropNone
        Me.Grid.Editable = flexEDNone

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = False
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = True
            'Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "E" Then
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            Me.GrdTBar.Buttons("Barcode").Enabled = False
            '  Me.GrdTBar.Buttons("Options").Enabled = True
            Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If

    ElseIf Me.TxtModFlag.Text = "N" Then

        If Not Me.LblTotalAll Is Nothing Then
            Me.LblTotalAll.Caption = 0
        End If

        If Not Me.txtTotal Is Nothing Then
            Me.txtTotal.Text = 0
        End If

        If Not Me.LblTotalQty Is Nothing Then
            Me.LblTotalQty.Caption = 0
        End If
    
        Me.Grid.OLEDropMode = flexOLEDropAutomatic
        Me.Grid.Editable = flexEDKbdMouse

        If Not Me.GrdTBar Is Nothing Then

            For i = 1 To Me.GrdTBar.Buttons.count
                Me.GrdTBar.Buttons(i).Enabled = True
            Next i

            '        Me.GrdTBar.Buttons("Barcode").Enabled = False
            'Me.GrdTBar.Buttons("Options").Enabled = True
            '        Me.GrdTBar.Buttons("GridCols").Enabled = True
        End If
    End If
    
End Sub

Private Sub GetMySetting()

    Dim StrSetting As String
    Dim StrShowSet As String
    Dim frmname As String

    Dim VarCols As Variant
    Dim VarColSet As Variant
    Dim i As Integer
    On Error Resume Next
    
    frmname = Me.Grid.Parent.Name
    StrSetting = GetSetting(SystemOptions.SysRegsAppPath, "Interface Setting" & "\" & user_id, frmname, "")

    If StrSetting = "" Then
        Exit Sub
    End If

    VarCols = Split(StrSetting, ";", , vbTextCompare)

    If UBound(VarCols) > 0 Then

        For i = 0 To UBound(VarCols)
            VarColSet = Empty
            VarColSet = Split(CStr(VarCols(i)), "-", , vbTextCompare)

            With Grid
                .ColPosition(.ColIndex(CStr(VarColSet(0)))) = CLng(VarColSet(1))
                
            End With

        Next i

    End If

    StrShowSet = GetSetting(SystemOptions.SysRegsAppPath, "Cols Setting" & "\" & user_id, frmname, "")

    If StrShowSet = "" Then
        Exit Sub
    End If

    VarCols = Split(StrShowSet, ";", , vbTextCompare)

    If UBound(VarCols) > 0 Then

        For i = 0 To UBound(VarCols)
            VarColSet = Empty
            VarColSet = Split(CStr(VarCols(i)), "-", , vbTextCompare)

            With Grid
                         If frmname = "FrmPO4" Then
                                     If VarColSet(0) = "Price" Then
                                    .ColHidden(.ColIndex(CStr(VarColSet(0)))) = True
                                    Else
                                    .ColHidden(.ColIndex(CStr(VarColSet(0)))) = CBool(VarColSet(1))
                                    End If
                        
                        Else
                           .ColHidden(.ColIndex(CStr(VarColSet(0)))) = CBool(VarColSet(1))
                           End If
            End With

        Next i

    End If

    StrSetting = ""
    'If Not TabFrm Is Nothing Then
    '    StrSetting = GetSetting(SystemOptions.SysRegsAppPath, "TabSetting" & "\" & User_ID & "\" & FrmName, SystemOptions.UserInterface, "")
    '    If StrSetting <> "" Then
    '        VarCols = Split(StrSetting, "-", , vbTextCompare)
    '        If SystemOptions.UserInterface = Val(VarCols(2)) Then
    '            TabFrm.Position = Val(VarCols(0))
    '            TabFrm.PicturePos = Val(VarCols(1))
    '        End If
    '    End If
    'End If
End Sub

Private Sub ShowGridColFrm()
    Dim IntColIndex As Integer
    Dim i As Integer

    Load FrmGridColsShow

    With FrmGridColsShow.FG
        .Rows = .FixedRows

        '-----------------
                If Me.Grid.ColIndex("NewItem") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "’‰› ÃœÌœ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "New Item"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "NewItem"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NewItem"))
        End If
        
        
        If Me.Grid.ColIndex("Ser") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„”·”· «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Ser"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Ser"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Ser"))
        End If

        '-----------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·’‰›"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Code"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Code"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Code"))
        '------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«”„ «·’‰›"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Name  "
        End If
          
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Name"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Name"))

        '------------
        '«·ÊÕœ…
        If Me.Grid.ColIndex("UnitID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ÊÕœ… «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Unit"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "UnitID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("UnitID"))
        End If
        
        '«··Ê‰
        If Me.Grid.ColIndex("ColorID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "·Ê‰ «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Item Color  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ColorID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ColorID"))
        End If

        '---------------------
        '«·„ﬁ«”
        If Me.Grid.ColIndex("ItemSize") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„ﬁ«” «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Size"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemSize"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemSize"))
        End If

        '-------------------
        '«·›∆…
        If Me.Grid.ColIndex("ClassId") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "›—“ «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Item Category  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ClassId"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ClassId"))
        End If

        '„Ê«ﬁ⁄ «·Õ›Ÿ
        If Me.Grid.ColIndex("StoreIDLoc") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„Êﬁ⁄ «· Œ“Ì‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Store Location"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "StoreIDLoc"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StoreIDLoc"))
        End If

        '-------------------
                    If Me.Grid.ColIndex("StoreID2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·„Œ“‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Store"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "StoreID2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StoreID2"))
        End If


    '********************************
        '«·«Ã—«¡«  «· ’ÕÌÕÌ…
        If Me.Grid.ColIndex("CorrectionID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·«Ã—«¡«  «· ’ÕÌÕÌ…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Corrective action"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "CorrectionID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CorrectionID"))
        End If

        '-------------------
    
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì…"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Quantity"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Count"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Count"))
        '-------------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·”⁄—"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Price"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Price"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Price"))
        '-------------------
        
        '-------------------
        If Me.Grid.ColIndex("DiscountType") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·Œ’„"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Discount Type"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DiscountType"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountType"))
        End If

        '-------------------
        If Me.Grid.ColIndex("DiscountVal") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·Œ’„"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Discount Value"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DiscountVal"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountVal"))
        End If

        '-------------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬁÌ„…"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Value"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Valu"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Valu"))
        '-------------------

        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Item Status  "
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemCase"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCase"))
        '-------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "·Â ”Ì—Ì«·"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Have Serial"
        End If
        
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "HaveSerial"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("HaveSerial"))
        '---------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "≈÷«›… «·”Ì—Ì«·"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Add Serials"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AddSerials"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AddSerials"))
        '-------------------
        .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «·”Ì—Ì«·"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  Serial NO"
        End If

        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Serial"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Serial"))
        '-------------------


 
        
        
        If Me.Grid.ColIndex("guaranteeTime") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„œ… «·÷„«‰ »«·‘Â—"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "guarantee Time"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "guaranteeTime"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("guaranteeTime"))
        End If
        
        If Me.Grid.ColIndex("ItemType") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·’‰› "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ItemType "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemType"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemType"))
        End If
        
        If Me.Grid.ColIndex("order_no") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·ÿ·»ÌÂ"
            Else
            
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "order no"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "order_no"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("order_no"))
        End If
        
       
        If Me.Grid.ColIndex("OrderArrivalDate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "   «—ÌŒ Ê’Ê· «·‘Õ‰…   "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Order Arrival Date  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "OrderArrivalDate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OrderArrivalDate"))
        End If
        
        If Me.Grid.ColIndex("LineShahn") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„’«—Ì› «·‘Õ‰"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LineShahn"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LineShahn"))
        End If
        
        If Me.Grid.ColIndex("ItemCostPrice") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— «· ﬂ·›…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ItemCostPrice"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemCostPrice"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCostPrice"))
        End If
        
        
        If Me.Grid.ColIndex("CostPriceTk") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— «· ﬂ·›… ··«” ·«„"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "CostPriceTk"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "CostPriceTk"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CostPriceTk"))
        End If
        
'CostPriceTk

        If Me.Grid.ColIndex("PofTransID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ „Œ’’"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "PofTransID"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "PofTransID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PofTransID"))
        End If
        
        If Me.Grid.ColIndex("ItemProfit") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—»Õ «·’ ›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ItemProfit"
            
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemProfit"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemProfit"))
        End If
      
        If Me.Grid.ColIndex("FoxyNo") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·»—‰«„Ã"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Prog No"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "FoxyNo"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("FoxyNo"))
        End If
                                        
        If Me.Grid.ColIndex("Account_Serial") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
            
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ﬂÊœ  «·Õ”«»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Account Serial"

            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Account_Serial"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Account_Serial"))
        End If
        
        If Me.Grid.ColIndex("AccountCode") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "code 2"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ﬂÊœ   2"

            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AccountCode"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AccountCode"))
        End If
        
        If Me.Grid.ColIndex("AccountName") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " «”„ «·Õ”«»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " AccountName"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AccountName"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AccountName"))
        End If
        
        If Me.Grid.ColIndex("GardQty") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " «·ﬂ„ÌÂ «·œ› —ÌÂ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Ac QTy  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GardQty"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GardQty"))
        End If
        
        If Me.Grid.ColIndex("Gardresult") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·›—ﬁ  "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Different  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Gardresult"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult"))
        End If
        
        If Me.Grid.ColIndex("Gardresult1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄Ã“  "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Minus  "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Gardresult1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult1"))
        End If
        
        If Me.Grid.ColIndex("Gardresult2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·“Ì«œ…  "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Increase"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Gardresult2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult2"))
        End If
        
        If Me.Grid.ColIndex("ProductionDate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «—ÌŒ   «·«‰ «Ã"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Production Date"
            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ProductionDate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ProductionDate"))
        End If

        If Me.Grid.ColIndex("GranteeData") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "»Ì«‰«  «·÷„«‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Gurantee Date"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GranteeData"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GranteeData"))
        End If
        
                If Me.Grid.ColIndex("GoldData") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ›«’Ì· «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Details"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GoldData"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GoldData"))
        End If
        
        
        
        
                  If Me.Grid.ColIndex("ParrtNoCode") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " part Code"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ParrtNoCode"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ParrtNoCode"))
        End If
        
        
                          If Me.Grid.ColIndex("ItemDetailedCode") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬂÊœ   Õ·Ì·Ì «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Detailed part Code"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemDetailedCode"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemDetailedCode"))
        End If
        
   
   
           
                          If Me.Grid.ColIndex("PrintName") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«”„ «·’‰› «·„»«⁄"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Print Name"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "PrintName"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PrintName"))
        End If
        
     
        'ItemCostPrice
        
        
        If Me.Grid.ColIndex("ExpiryDate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " «—ÌŒ   «·«‰ Â«¡"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Expire Date"
            End If
           
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ExpiryDate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExpiryDate"))
        End If
        
        
                       If Me.Grid.ColIndex("IsExpirDate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " «—ÌŒ ÌœÊÌ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Date Manual"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "IsExpirDate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("IsExpirDate"))
        End If




        'LotNO
        If Me.Grid.ColIndex("LotNO") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ «··Êÿ"
 
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LOT NO"

            End If
            
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LotNO"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LotNO"))
        End If
        
        If Me.Grid.ColIndex("DamageshowQty") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·Â«·ﬂ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Damage"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DamageshowQty"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DamageshowQty"))
        End If
          
        If Me.Grid.ColIndex("GranteeData") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "»Ì«‰«  «·÷„«‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Gurantee Date"
            End If
           
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GranteeData"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GranteeData"))
        End If

       If Me.Grid.ColIndex("GoldData") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ›«’Ì· «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Details"
            End If
           
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GoldData"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GoldData"))
        End If

       If Me.Grid.ColIndex("ShowQty2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Ê“‰ «·–Â»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Gold Weight"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ShowQty2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowQty2"))
        End If
      
      
             If Me.Grid.ColIndex("showPrice2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— «·Ã—«„"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Gram Value"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "showPrice2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("showPrice2"))
        End If
        
        
                     If Me.Grid.ColIndex("Value2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·–Â»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Gold Value"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Value2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Value2"))
        End If
        
        
        
      

              If Me.Grid.ColIndex("ShowQty1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Ê“‰ «·«·„«”"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Diamond Weight"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ShowQty1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowQty1"))
        End If
        
        
                If Me.Grid.ColIndex("showPrice1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— «·ﬁÌ—«ÿ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Carat Price"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "showPrice1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("showPrice1"))
        End If
        
        
              
                If Me.Grid.ColIndex("Value1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·√·„«”"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Diamond  Total"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Value1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Value1"))
        End If
        
        '*********************************
                 '⁄„·… «·«·„«”
                    If Me.Grid.ColIndex("ScurrencyID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄„·…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Currency"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ScurrencyID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ScurrencyID"))
        End If
        '*********************************
        
              '*********************************
                 '„⁄œ· ⁄„·… «·«·„«”
                    If Me.Grid.ColIndex("Scurrenyrate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„⁄œ·"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Rate"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Scurrenyrate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Scurrenyrate"))
        End If
        '*********************************
        
                      '*********************************
                 '  —ﬁ„ «·«„— ⁄·Ï „”‰ÊÏ «·”ÿ— ›Ì ”‰œ «” ·«„ «‰ «Ã  «„
                    If Me.Grid.ColIndex("NProductionOrderNO") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ «„— «·«‰ «Ã"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Production Order  NO"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "NProductionOrderNO"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NProductionOrderNO"))
        End If
        '*********************************
        
        
        
     '//////////////////////////////////////////////////////////////////////////////////////////
             '*********************************
        
                      '*********************************
                 '  ÿ·» «·‘—«¡ «·œ«Œ·Ì
                    If Me.Grid.ColIndex("LastPurchaseDate") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «—ÌŒ «Œ— ‘—«¡"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Last Purchase Date"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LastPurchaseDate"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchaseDate"))
        End If
        
       
                    If Me.Grid.ColIndex("LastPurchasePrice") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«Œ—  ”⁄—   ‘—«¡"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Last Purchase Price"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LastPurchasePrice"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchasePrice"))
        End If
               
               
               
               
                    If Me.Grid.ColIndex("LastPurchaseqty") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ﬂ„Ì… «Œ— ‘—«¡"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Last Purchase Qty"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LastPurchaseqty"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchaseqty"))
        End If
               
               
                                 If Me.Grid.ColIndex("ItemBalance") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—’Ìœ «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Balance "
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemBalance"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemBalance"))
        End If
        
        
           If Me.Grid.ColIndex("DeliveredQty2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·„”·„…  "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Delivered Qty"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DeliveredQty2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DeliveredQty2"))
        End If
        
        
        
                   If Me.Grid.ColIndex("RecivePeriod") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„œ… «· Ê—Ìœ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " Recive Period"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RecivePeriod"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RecivePeriod"))
        End If
        
        
               
                    If Me.Grid.ColIndex("AverageIssue") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„ Ê”ÿ «·«”‰Â·«ﬂ «·‘Â—Ì"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "íMonthly Average"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AverageIssue"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AverageIssue"))
        End If
        
                            If Me.Grid.ColIndex("AverageIssueyraly") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «·«”‰Â·«ﬂ «·”‰ÊÌ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Yearly Average"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AverageIssueyraly"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AverageIssueyraly"))
        End If
        
        
        
                          If Me.Grid.ColIndex("RequestLimit") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Õœ «⁄«œ… «·ÿ·»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "íRequest Limit"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RequestLimit"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RequestLimit"))
        End If
        
       '//////////////////////////////////////////////////////////////////////////////////////////
       
     '*********************
     
        
                          If Me.Grid.ColIndex("Shipping") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·‘Õ‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Shipping"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Shipping"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Shipping"))
        End If
     
                          If Me.Grid.ColIndex("countris") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«· ’‰Ì⁄"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Made In"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "countris"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("countris"))
        End If
        
     
     '***********************
        
             
        '***************************
       If Me.Grid.ColIndex("catalog") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "   «·ﬂ «·ÊÃ "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " catalog"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "catalog"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("catalog"))
        End If
        '**************************
        
        
        
        
        
        
        
             If Me.Grid.ColIndex("Salaries1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«ÃÊ— ’Ì«€Â"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Formulation fees"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Salaries1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Salaries1"))
        End If
        
        
                
             If Me.Grid.ColIndex("Salaries2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«ÃÊ—  —ﬂÌ»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Installation fees"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Salaries2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Salaries2"))
        End If
        
        
          If Me.Grid.ColIndex("TotalCost") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·«Ã„«·Ì «·⁄«„"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Total Cost"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "TotalCost"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalCost"))
        End If
        
        
        '---------------------
        
        
        '  ««·„Ê—œÌ‰
        If Me.Grid.ColIndex("SupplierID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·„Ê—œÌ‰"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Supplier Name"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "SupplierID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SupplierID"))
        End If

        '-------------------
        
       '-------------------
        If Me.Grid.ColIndex("ExtraType") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·⁄„Ê·…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "commision Type"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ExtraType"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExtraType"))
        End If

        '-------------------
        If Me.Grid.ColIndex("ExtraVal") <> -1 Then '
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·⁄„Ê·…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "commision Value"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ExtraVal"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExtraVal"))
        End If

 


        '-------------------
                If Me.Grid.ColIndex("SBillNO") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·›« Ê—…"
            Else
            
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Bill no"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "SBillNO"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SBillNO"))
        End If
        
        
                         If Me.Grid.ColIndex("ItemCostPrice") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«· ﬂ·›…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Cost Price"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemCostPrice"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCostPrice"))
        End If
            
         
             
        
         '??????
   
             ' Exit Sub
                          If Me.Grid.ColIndex("LAXIS") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LAXIS"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LAXIS"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LAXIS"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LAXIS"))
        End If
        
        
        
        
                          If Me.Grid.ColIndex("LCYL") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LCYL"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LCYL"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LCYL"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LCYL"))
        End If
         
         
        
                          If Me.Grid.ColIndex("LSPH") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LSPH"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LSPH"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LSPH"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LSPH"))
        End If
        
        
        
                          If Me.Grid.ColIndex("RAXIS") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RAXIS"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RAXIS"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RAXIS"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RAXIS"))
        End If
                
                
                
                         If Me.Grid.ColIndex("RCYL") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RCYL"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RCYL"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RCYL"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RCYL"))
        End If
        
        
        
                
                         If Me.Grid.ColIndex("RSPH") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RSPH"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RSPH"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RSPH"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RSPH"))
        End If
        
        
        
  '*****************
              If Me.Grid.ColIndex("Height") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·«— ›«⁄"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Height"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Height"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Height"))
        End If
        ''''«·⁄—÷
             If Me.Grid.ColIndex("Width") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄—÷"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Width"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Width"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Width"))
        End If
           
       '
         
               ''''«·„”«Õ…
             If Me.Grid.ColIndex("Area") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·„”«Õ…"
                            Else
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Area"
                            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Area"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Area"))
          End If
                     ''''«·⁄œœ
             If Me.Grid.ColIndex("NoCount") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄œœ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "No"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "NoCount"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NoCount"))
        End If

'*****************
        ''//////////  «Ê  ÕÃ„ „»Ì⁄«  «·⁄«„
             If Me.Grid.ColIndex("ItemSalesYear") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Status"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemSalesYear"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemSalesYear"))
        End If




             If Me.Grid.ColIndex("Remarks") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„·«ÕŸ« "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Remarks"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Remarks"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Remarks"))
        End If





   '*****************
        ''//////////ÃœÌœ «Ê „ œ«Ê·
             If Me.Grid.ColIndex("New_Rolling") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Item Status"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "New_Rolling"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("New_Rolling"))
        End If


        
                        '„ »ﬁÌ Õœ «·ÿ·»
        If Me.Grid.ColIndex("ReminRequ") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„ »ﬁÌ Õœ «·ÿ·»"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Remaining Request"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ReminRequ"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ReminRequ"))
        End If


                   If Me.Grid.ColIndex("Select") <> -1 Then
             .Rows = .Rows + 1
        i = i + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ÕœÌœ"
        Else
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Select"
        End If
        
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Select"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Select"))
        
        End If

'--------------------------------------------------------------------

                  If Me.Grid.ColIndex("L") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "L"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "L"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "L"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("L"))
        End If
                '''/////////////
          If Me.Grid.ColIndex("TotalPrice") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·”⁄— «·‘«„·"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Total"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "TotalPrice"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalPrice"))
        End If
        
        '''/////////////
          If Me.Grid.ColIndex("W") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "W"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "W"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "W"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("W"))
        End If
         If Me.Grid.ColIndex("H1") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "H1"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "H1"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "H1"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("H1"))
        End If
      ''///
            If Me.Grid.ColIndex("H2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "H2"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "H2"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "H2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("H2"))
        End If
        '--------------------------------------------------------------------
                            If Me.Grid.ColIndex("Vatyo") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰”»… «· VAT % "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "VAT %"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Vatyo"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Vatyo"))
        End If
        
        
        
                    If Me.Grid.ColIndex("Vat") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «· VAT"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "VAT  "
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Vat"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Vat"))
        End If
        
        
                
        If Me.Grid.ColIndex("TypeVAT") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «· VAT "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "VAT Type"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "TypeVAT"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TypeVAT"))
        End If
        
        
             If Me.Grid.ColIndex("MintName") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ‰Ê⁄ «·’Ì«‰…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Maintenance"
            End If
           
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "MintName"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MintName"))
          End If
                If Me.Grid.ColIndex("GroupMint") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "›∆… «·’Ì«‰…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Group"
            End If
           
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "GroupMint"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GroupMint"))
        End If


         If Me.Grid.ColIndex("QtyFaqtors") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄·«ﬁ… „⁄ «·ÊÕœ…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Qty Factor"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "QtyFaqtors"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("QtyFaqtors"))
        End If


'*********************************************
If Me.Grid.ColIndex("MixNo") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ﬂÊœ «·„ﬂ”"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Mix. No"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "MixNo"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MixNo"))
        End If

'*********************************************

'*******************

          If Me.Grid.ColIndex("MaxUnitID") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «·ÊÕœ… «·ﬂ»—Ï"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Max Unit"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "MaxUnitID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MaxUnitID"))
        End If
       ''//////////
          If Me.Grid.ColIndex("MaxQty") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «·ﬂ„Ì… «·ﬂ»—Ï"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Max Qty"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "MaxQty"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MaxQty"))
        End If

'****************
   If Me.Grid.ColIndex("CusID2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «”„ «·⁄„Ì·"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Customer Name"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "CusID2"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CusID2"))
        End If
'***************************************************************
       If Me.Grid.ColIndex("NoHours") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "⁄œœ «·”«⁄« "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "No. Hours"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "NoHours"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NoHours"))
        End If
        
         If Me.Grid.ColIndex("PriceNoHours") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ﬂ·›… «·”«⁄… "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Cost Hour"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "PriceNoHours"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PriceNoHours"))
        End If
        
          If Me.Grid.ColIndex("TotalPriceNoHours") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " «Ã„«·Ì  ﬂ·›… «·”«⁄«  "
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Total Cost Hours"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "TotalPriceNoHours"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalPriceNoHours"))
        End If
 
          If Me.Grid.ColIndex("TotalInvoiceQty") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·«Ã„«·Ì…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Total Û QTY"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "TotalInvoiceQty"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalInvoiceQty"))
        End If


          If Me.Grid.ColIndex("ISSUEDQTY") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "  «·„‰’—› ”«»ﬁ«"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ISSUED Û QTY"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ISSUEDQTY"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ISSUEDQTY"))
        End If



 ShowGridColFrm1 (i)
        
    End With


End Sub
Sub ShowGridColFrm1(i As Integer)
  With FrmGridColsShow.FG
  
'**************************
       
        '«·ÿÊ·
           If Me.Grid.ColIndex("length") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ÿÊ·"
                            Else
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "length"
                            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "length"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("length"))
          End If
       
       
               ''''ﬁÿ— Œ«—ÃÌ
             If Me.Grid.ColIndex("OUTR") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÿ— Œ«—ÃÌ"
                            Else
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "OUTR"
                            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "OUTR"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OUTR"))
          End If
       
       
       '***************************
        
     '**************************
       
               ''''ﬁÿ— œ«Œ·Ì
             If Me.Grid.ColIndex("INR") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÿ— œ«Œ·Ì"
                            Else
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "INR"
                            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "INR"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("INR"))
          End If
       
       '***************************
      
       '**************************
       
               ''''  ⁄—÷ «·„—›ﬁ« 
             If Me.Grid.ColIndex("ShowAttatch2") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " ⁄—÷ «·„—›ﬁ«    "
                            Else
                                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Show Attatch "
                            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "INR"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowAttatch2"))
          End If
       
       '***************************
           
           
  
  
            If Me.Grid.ColIndex("ManufactCompany") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·‘—ﬂ… «·’«‰⁄…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Manufacturing Company"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ManufactCompany"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ManufactCompany"))
        End If
        
           If Me.Grid.ColIndex("CountryOrg") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = " »·œ «·„‰‘√"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Country of Origin"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "CountryOrg"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CountryOrg"))
        End If
  
  
          If Me.Grid.ColIndex("EmpID4") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "   «”„ «·„ÊŸ›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Employee Name"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "EmpID4"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("EmpID4"))
        End If
        
  
  
           If Me.Grid.ColIndex("SalesValue") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— »Ì⁄ «·’‰›"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Sales Value"
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "SalesValue"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SalesValue"))
        End If



  
           If Me.Grid.ColIndex("UnitPrice") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "”⁄— «·ÊÕœ…  " ' «·«”„ ⁄—»Ì
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Unit price" ' «·«”„ «‰Ã·Ì“Ì
            End If
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "UnitPrice"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("UnitPrice"))
        End If

'***************************************************GLASSES'***************************************************
        If Me.Grid.ColIndex("LIPD") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LIPD"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LIPD"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LIPD"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LIPD"))
        End If
        
                  If Me.Grid.ColIndex("RIPD") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RIPD"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RIPD"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RIPD"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RIPD"))
        End If
        
        
                  If Me.Grid.ColIndex("LADD") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LADD"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LADD"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LADD"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LADD"))
        End If
        
                  If Me.Grid.ColIndex("RADD") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RADD"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RADD"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RADD"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RADD"))
        End If
        
                  If Me.Grid.ColIndex("LSH") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LSH"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LSH"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LSH"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LSH"))
        End If
        
         
                           If Me.Grid.ColIndex("RSH") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RSH"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RSH"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RSH"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RSH"))
        End If
        
        
                                If Me.Grid.ColIndex("LPRISM") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LPRISM"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "LPRISM"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "LPRISM"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LPRISM"))
        End If
        
        
                                If Me.Grid.ColIndex("RPRISM") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RPRISM"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "RPRISM"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "RPRISM"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RPRISM"))
        End If
  
  
                                If Me.Grid.ColIndex("project") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·„‘—Ê⁄"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "project"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "project"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("project"))
        End If
    
    
                                 If Me.Grid.ColIndex("pand") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·»‰œ"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "pand"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "pand"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("pand"))
        End If
    
    
        
    
                                 If Me.Grid.ColIndex("opera") <> -1 Then
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·⁄„·Ì…"
            Else
                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "opera"
            End If

            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "opera"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("opera"))
        End If
    
    
    
    
    
  '***************************************************GLASSES'***************************************************
  
  .AutoSize 0, .Cols - 1, False
  
  End With
  
      Set FrmGridColsShow.FgGrid = Me.Grid
    FrmGridColsShow.show vbModal
End Sub
Public Sub SaveMySetting()
    Dim i As Integer
    Dim StrTemp As String
    Dim StrShow As String
    Dim frmname As String
    frmname = Me.Grid.Parent.Name

  '  If frmname = "Projects" Then Exit Sub
                              
    For i = 0 To Grid.Cols - 1
        StrTemp = StrTemp & Grid.ColKey(i) & "-" & i & ";"
        StrShow = StrShow & Grid.ColKey(i) & "-" & Grid.ColHidden(i) & ";"
    Next i

    StrTemp = Trim(StrTemp)
    StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
    StrShow = Trim(StrShow)
    StrShow = mId(StrShow, 1, Len(StrShow) - 1)
    SaveSetting SystemOptions.SysRegsAppPath, "Interface Setting" & "\" & user_id, frmname, StrTemp
    SaveSetting SystemOptions.SysRegsAppPath, "Cols Setting" & "\" & user_id, frmname, StrShow

    '-----------------------------------------

End Sub


Private Sub TxtPrice_GotFocus()


    TxtPrice.backcolor = vbYellow
    SelectText TxtPrice

    If Me.DCboItemName.BoundText <> "" Then
        If m_ShowQtyToolTip = True Then
        'Wael Updated
           If m_GridTrans2 <> InvoiceTransactionCompose Then
                ShowItemTip ItemPriceToolTip, val(Me.DCboItemName.BoundText), 0, 0, , False, Me.TxtPrice.hwnd
            End If
        End If
    End If
If SystemOptions.UserInvoiceChangePrice = 0 And Me.GridTrans = InvoiceTransaction Then
 TxtPrice.locked = True
End If
End Sub

 
   Private Sub TxtPrice_KeyDown(KeyCode As Integer, _
                             Shift As Integer)

    
    
    If KeyCode = vbKeyReturn Then
        CmdAddData_Click
        'CmdAddData.SetFocus
        
        If val(TxtQuantity.Text) = 0 Then
            Exit Sub
        End If
        If Not TxtCustCode Is Nothing Then
            TxtCustCode.SetFocus
        Else
            DCboItemCode.SetFocus
        End If
    
        
        
                                            
              If SystemOptions.AllowItemsShortName = True Then
                                    If Not TxtShortName Is Nothing Then
                                        TxtShortName.SetFocus
                            End If
           End If

    End If

    Exit Sub
ErrTrap:
End Sub


Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPrice.Text, 0)
End Sub

Private Sub TxtPrice_LostFocus()
    'TxtPrice.BackColor = vbYellow
    TxtPrice.backcolor = vbWhite
    HideComment
End Sub

Private Sub TxtQuantity_GotFocus()
    TxtQuantity.backcolor = vbYellow
End Sub

Private Sub TxtQuantity_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
   
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.Text, 0)
End Sub

 
 Private Sub TxtQuantity_LostFocus()
'Wael  Update
Dim Msg As String
    TxtQuantity.backcolor = vbWhite
        If Me.GridTrans2 = InvoiceTransactionCompose Then
             If val(TxtQuantity.Text) = 0 Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                                 Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰›"
                             Else
                             Msg = "Enter item Qty"
                             End If
                             
        '     Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
             MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             TxtQuantity.SetFocus
             Exit Sub
            End If
        End If
End Sub


Private Sub AddItemPrice(LngItemID As Long, _
                         Col As Long, _
                         Row As Long, _
                         Optional LngUnitID As Long, _
                         Optional Purchase As Long)

    With Grid

        If TxtModFlag.Text = "N" Or TxtModFlag.Text = "E" Then
            If Col = .ColIndex("unitId") Or Col = .ColIndex("Code") Or Col = .ColIndex("Name") Then
                .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID, Purchase)
            

                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                    .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName)
                    If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                        .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID, Purchase)
                    End If
                 
                 
                End If
                
       
                
                If Me.GridTrans = Returntransaction Then
                
             
                    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                
                    Fromdate = FirstPeriodDateInthisYear
                    GetlastPurchasedata 20, (LngItemID), Fromdate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                
                  .TextMatrix(Row, .ColIndex("Price")) = LastPurchasePrice

                 End If
                checkQtybylengthoption Row, LngItemID
                
      
            
            End If
        End If

    End With

End Sub
Public Function CalculateInvProfit(LngRow As Long) As Double
    Dim DblItemProfit As Double
    Dim DblItemCost As Double
    Dim DblTotalItemCost As Double
    Dim LngItemID As Long
    Dim IntHaveSerial As Integer
    Dim DblInvProfit As Double
    Dim StrItemSerial As String
    Dim StrTransID As String
    Dim LngUnitID As Long
   On Error GoTo ErrTrap

    If Me.TxtModFlag.Text <> "N" And Me.TxtModFlag.Text <> "E" Then
        Exit Function
    End If

    If Me.Grid.Rows = Me.Grid.FixedRows Then Exit Function

    With Me.Grid
        LngItemID = val(Me.Grid.TextMatrix(LngRow, Grid.ColIndex("Code")))
        LngUnitID = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))

        If LngItemID <> 0 Then

            Select Case Me.Grid.Cell(flexcpChecked, LngRow, Me.Grid.ColIndex("HaveSerial"))

                Case flexChecked
                    IntHaveSerial = 1

                Case flexUnchecked
                    IntHaveSerial = 2
            End Select

            StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))

            If Not (IntHaveSerial = 1 And StrItemSerial = "") Then
           If Me.GridTrans <> InvoiceTransaction And Me.GridTrans <> Returntransaction Then Exit Function
                If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
                    'DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID, ModernWeightAverage, , , , , LngUnitID)
                    
                    DblItemCost = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrItemSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(Me.TxtNots.Text), LngUnitID, val(Me.StoreName.BoundText))
                
                Else
                    
                    DblItemCost = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrItemSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(Me.TxtNots.Text), LngUnitID, val(Me.StoreName.BoundText))
                    '  DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID)
                End If
          
            
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemCostPrice")) = DblItemCost
            
                If Me.Grid.ColIndex("ItemProfit") = -1 Then Exit Function
                DblTotalItemCost = (DblItemCost * val(Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("Count"))))
                
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemProfit")) = Round(val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) - DblTotalItemCost, SystemOptions.SysDefCurrencyForamt)
            
                
                '------------------------------------------------------------------
           If SystemOptions.showcostColorininvoice = True Then
                If val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) = DblTotalItemCost Then
                    Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbYellow
                ElseIf val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) < DblTotalItemCost Then
                    Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbRed
                Else
                    Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = 0
                End If
            Else
                     Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = 0
            End If
            
                '------------------------------------------------------------------
            
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("PofTransID")) = StrTransID

                If Not LblInvProfit Is Nothing Then
                    DblInvProfit = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemProfit"), .Rows - 1, .ColIndex("ItemProfit"))
                    Me.LblInvProfit.Caption = DblInvProfit

                    If DblInvProfit > 0 Then
                        Me.LblInvProfit.ForeColor = &H4000&
                    ElseIf DblInvProfit = 0 Then
                        Me.LblInvProfit.ForeColor = vbBlack
                    ElseIf DblInvProfit < 0 Then
                        Me.LblInvProfit.ForeColor = vbRed
                    End If
                End If
            End If
        End If

    End With

    CalculateInvProfit = DblInvProfit
    ColorExperDate
    Exit Function
ErrTrap:
End Function


 


   
 Public Function AddItemInGrid2(LngRow As Long, _
                              LngItemID As Long, _
                              Optional DblQty As Double = 1, _
                              Optional SngItemPrice As Single = 0, _
                              Optional IntItemCase As Integer = 0, _
                              Optional StrItemSerial As String = "", _
                              Optional FromHeader As Boolean = False, _
                              Optional DefalutUnitID As Long = 0, _
                              Optional DefalutUnitName As String, Optional freeitemsqty As Integer = 0, _
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer, Optional StoreId As Integer, Optional lotNo As String, Optional ExpiryDate As String, Optional ProductionDate As String)

    Dim RsParts As ADODB.Recordset
    Dim StrSQL  As String
    Dim i As Long, j As Long
    Dim DblPartQty As Double
    Dim longunitid As Long
    Dim UnitName As String
    Dim LngPartItemIndex As Long
    Dim LngNewPartRow As Long
    Dim LngAddCount As Long
    Dim StrSerial As String
    Dim Frm As FrmAddItemAttachs
    Dim LngAttachedItemID As Long
    Dim DblAttachedQty As Double
    Dim DblAttachedPrice As Single
    Dim RsUnitData As ADODB.Recordset
              
                           
                               
    If LngRow = 0 Then
        LngRow = Me.NewEmptyRow
    ElseIf LngRow > Me.Grid.Rows - 1 Then
        Me.Grid.Rows = LngRow + 1
    End If

    If Not (RsItems.BOF Or RsItems.EOF) Then
        If RsItems("ItemID").value <> LngItemID Then
            RsItems.find "ItemID=" & LngItemID & "", , adSearchForward, 1
        End If
    End If

    If Not (RsItems.BOF Or RsItems.EOF) Then

        'C???? C?????
        With Me.Grid
  
            If RsItems("AssbliedItem").value = True Then
                '   StrSQL = "SELECT  unitid,TableID, ItemID, PartItemID, PartItemQty, PartItemPrice"
                '   StrSQL = StrSQL + " FROM TblItemsParts Where ItemID=" & LngItemID
                '   StrSQL = StrSQL + " Order By TableID"
            
                StrSQL = "SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + " dbo.TblItemsParts.ItemID , dbo.TblItemsParts.TableID, dbo.TblUnites.unitname, dbo.TblUnites.UnitNamee"
                StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & LngItemID & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
                Set RsParts = New ADODB.Recordset
                RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsParts.EOF Or RsParts.BOF) Then

                    For i = 0 To RsParts.RecordCount - 1
                        LngPartItemIndex = LngPartItemIndex + 1
                        '--------------------------------------
                        'Add to grid the New Parts Rows
                        '--------------------------------------
 
                        longunitid = IIf(IsNull(RsParts("UnitID").value), 0, RsParts("UnitID").value)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            UnitName = IIf(IsNull(RsParts("unitname").value), 0, RsParts("unitname").value)
                        Else
                            UnitName = IIf(IsNull(RsParts("unitnamee").value), 0, RsParts("unitnamee").value)
                        End If

                        DblPartQty = IIf(IsNull(RsParts("PartItemQty").value), 0, RsParts("PartItemQty").value)
                        DblPartQty = DblPartQty * DblQty
                        LngNewPartRow = LngRow + (LngPartItemIndex - 1)
                        Me.Grid.AddItem "", LngNewPartRow
                        RsItems.MoveFirst
                        AddItemInGrid LngNewPartRow, RsParts("PartItemID").value, DblPartQty, , , , , longunitid, UnitName
                        RsParts.MoveNext
                    Next i

                End If
            
                '--------??C?E C????C? C?????E EC????
                If Not (RsItems.BOF Or RsItems.EOF) Then
                    If RsItems("ItemID").value <> LngItemID Then
                        RsItems.find "ItemID=" & LngItemID & "", , adSearchForward, 1
                    End If
                End If

                If RsItems("RelatedItem").value = True Then
                    Set Frm = New FrmAddItemAttachs
                    Frm.LoadItemData LngItemID, DblQty
                    Frm.show vbModal

                    If Frm.UserCancelled = True Then
                        Unload Frm
                    Else

                        For j = Frm.FG.FixedRows To Frm.FG.Rows - 1

                            If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                                LngAttachedItemID = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                                DblAttachedQty = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                                DblAttachedPrice = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                        AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice, , , , , , , , , , , , , , 4
             
                         
                            
                            End If

                        Next j

                        Unload Frm
                    End If
                End If

                Exit Function
            End If
             .TextMatrix(LngRow, .ColIndex("LotNO")) = lotNo
             .TextMatrix(LngRow, .ColIndex("ExpiryDate")) = ExpiryDate
             .TextMatrix(LngRow, .ColIndex("ProductionDate")) = ProductionDate
             .TextMatrix(LngRow, .ColIndex("Count")) = DblQty
             .TextMatrix(LngRow, .ColIndex("Code")) = LngItemID
             .TextMatrix(LngRow, .ColIndex("Name")) = LngItemID
             .TextMatrix(LngRow, .ColIndex("HaveSerial")) = RsItems("HaveSerial").value
                           
    If .ColIndex("showPrice2") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("showPrice2")) = GetDefaultGoldPrice
            End If
            

            If .ColIndex("guaranteeTime") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("guaranteeTime")) = IIf(IsNull(RsItems("GuaranteeValue").value), 0, RsItems("GuaranteeValue").value)
            End If

            'DblQty = Val(.TextMatrix(LngRow, .ColIndex("Count")))
            DblQty = IIf(DblQty = 0, 1, DblQty)
            .TextMatrix(LngRow, .ColIndex("Count")) = DblQty
            .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1 'set the defalut item case to new case

  If Me.Grid.ColIndex("ItemType") <> -1 Then
   .TextMatrix(LngRow, .ColIndex("ItemType")) = RsItems("ItemType").value
  End If
            If .ColIndex("DiscountType") <> -1 Then
            
            If freeitemsqty = 0 Then
                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
             Else
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
             
         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                    .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                    
                    
             End If
            End If

      If .ColIndex("ExtraType") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraType")) = 1
            End If
            
            If .ColIndex("order_no") <> -1 Then
                ' .TextMatrix(LngRow, .ColIndex("order_no")) = GetItemOrderNo(LngItemID)
            End If
             .TextMatrix(LngRow, .ColIndex("SalesValue")) = GetItemPrice(LngItemID, 1)
            If SngItemPrice = 0 Then
                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Then
                    'Put Sales Price
                    .TextMatrix(LngRow, .ColIndex("Price")) = GetItemPrice(LngItemID, 1)
                ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                    '?? ?C?E OCOE C?E???CE ??? ??? ??? C?E???E
                    StrSerial = Trim$(.TextMatrix(LngRow, .ColIndex("Serial")))
                    Dim UnitID As Long
 
                    UnitID = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                    .TextMatrix(LngRow, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), UnitID, val(Me.StoreName.BoundText))
                
                
                ElseIf Me.GridTrans = ProductionOrder Then
                    .TextMatrix(LngRow, .ColIndex("Price")) = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
          
                
            ElseIf Me.GridTrans = internalorder Then
                
                
               
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    Fromdate = FirstPeriodDateInthisYear
 
    
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Or Grid.ColIndex("LastPurchaseqty") <> -1 Or Grid.ColIndex("LastPurchasePrice") <> -1 Or Grid.ColIndex("RequestLimit") <> -1 Then
                           
                           GetlastPurchasedata 20, (LngItemID), Fromdate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                           End If
                           
                           
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchaseDate")) = LastPurchaseDate
                           End If
                           
                                If Grid.ColIndex("LastPurchasePrice") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchasePrice")) = LastPurchasePrice
                           End If
                           
                    If Grid.ColIndex("LastPurchaseqty") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchaseqty")) = LastPurchaseqty
                           End If
                           
                               If Grid.ColIndex("RequestLimit") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("RequestLimit")) = IIf(IsNull(RsItems("RequestLimit").value), 0, RsItems("RequestLimit").value)
                           End If
                           
                           
                            If Grid.ColIndex("AverageIssue") <> -1 Then
                            GetMonthlyAverage 19, (LngItemID), Fromdate, Date, AverageMonthly
                               .TextMatrix(LngRow, .ColIndex("AverageIssue")) = AverageMonthly
                           End If
                           
                           
                                   If Grid.ColIndex("AverageIssueyraly") <> -1 Then
                            GetYearlyAverage 19, (LngItemID), Fromdate, Date, AverageIssueyraly
                               .TextMatrix(LngRow, .ColIndex("AverageIssueyraly")) = AverageIssueyraly
                           End If
                           
                           
                  '     If Grid.ColIndex("ItemBalance") <> -1 Then
           '                 GetMonthlyAverage 19, (LngItemID), Fromdate, Date, ItemBalance
           '                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = ItemBalance
              
   ' Dim RsTest As ADODB.Recordset
   ' Dim DblOutQty As Double
   ' Dim LngColorID As Long
   ' Dim LngClassId As Long
'
'    Dim StrItemSize As String
'                                   LngColorID = ColorID ' val(.TextMatrix(LngRow, .ColIndex("ColorID")))
'                                StrItemSize = itemsize ' .TextMatrix(LngRow, .ColIndex("ItemSize"))
'                                LngClassId = ClassId ' val(.TextMatrix(LngRow, .ColIndex("ClassId")))
'
'                                If Trim(Grid.TextMatrix(LngRow, .ColIndex("Serial"))) <> "" Then
'                                    StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))
'                                Else
'                                    StrItemSerial = ""
'                                End If
'
'            Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), TransactionDate, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
'              If RsTest.EOF Or RsTest.BOF Then
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = 0
'              Else
'              ' (Round(RsTest("totalqty").value, 2)
'
'              ItemBalance = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
'
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = Round(ItemBalance, 2)
'              End If
              
                           
'                           End If
                           

                Else
                    'Put Purchase Price
                    .TextMatrix(LngRow, .ColIndex("Price")) = IIf(IsNull(RsItems("PurchasePrice").value), 0, RsItems("PurchasePrice").value)
                End If

            Else
                'Put the Sended Price
                .TextMatrix(LngRow, .ColIndex("Price")) = SngItemPrice
            End If

            '----------------------------------
            If IntItemCase = 0 Then
                If IsNull(RsItems("ItemCase").value) Then
                    .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1
                Else
                    .TextMatrix(LngRow, .ColIndex("ItemCase")) = RsItems("ItemCase").value
                End If

            Else
                .TextMatrix(LngRow, .ColIndex("ItemCase")) = IntItemCase
            End If

            .TextMatrix(LngRow, .ColIndex("Serial")) = StrItemSerial
            '        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            '            .TextMatrix(LngRow, .ColIndex("Serial")) = TxtSerial.text
            '        End If
        
            If Me.GridTrans <> OpeningBalance And Me.GridTrans <> Returntransaction And Me.GridTrans <> NewGard Then
           '     .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
                
                
                
                      If freeitemsqty = 0 Then
                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
                .TextMatrix(LngRow, .ColIndex("DiscountVal")) = 0
             Else
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
             
         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                    .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                    
                    
             End If
                
                If Me.Grid.ColIndex("ExtraType") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraType")) = 1
                End If
                
                If Me.Grid.ColIndex("ExtraVal") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraVal")) = 0
                End If
                If Me.Grid.ColIndex("ScurrencyID") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ScurrencyID")) = 1
                End If
                
            End If

            If Me.Grid.ColIndex("ItemType") <> -1 Then
                If Not (IsNull(RsItems("ItemType").value)) Then
                    .TextMatrix(LngRow, .ColIndex("ItemType")) = RsItems("ItemType").value

                    If val(.TextMatrix(LngRow, .ColIndex("ItemType"))) = 1 Then
                                            .Cell(flexcpPicture, LngRow, .ColIndex("Ser")) = _
                                                 mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
                    Else

                        If .ColIndex("Ser") <> -1 Then
                            .Cell(flexcpPicture, LngRow, .ColIndex("Ser")) = Nothing
                        End If
                    End If

                Else
                    .TextMatrix(LngRow, .ColIndex("ItemType")) = 0
                End If
            End If

            '-----------------------------------
            .TextMatrix(LngRow, .ColIndex("ColorID")) = ColorID
            .TextMatrix(LngRow, .ColIndex("ItemSize")) = itemsize
            .TextMatrix(LngRow, .ColIndex("Classid")) = ClassId
       '     .TextMatrix(LngRow, .ColIndex("StoreID")) = 1


            
                          'C???? C????? UC?EC ??C??
                        If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) And discountype = 4 Then
                                If .ColIndex("DiscountType") <> -1 Then
                                        If SystemOptions.attacheditemsisfree = True Then
                                                                 .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                                                                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                                                   
                                        End If
                                End If
                                
                        End If
                        
'???? C?C??C?
Dim discId As Double
Dim discValue As Double
                        If (Me.GridTrans = InvoiceTransaction) Then
                                            If CheckItem(LngItemID, DtpBillDate, discId, discValue, , val(Branch.BoundText)) = True Then
 
                                                                      If .ColIndex("DiscountType") <> -1 Then
                                                                      
                                                                   '0 value 1 percentage
                                                                If discId = 0 Then
                                                                discId = 2
                                                                ElseIf discId = 1 Then
                                                                discId = 3
                                                                             ElseIf discId = 4 Then
                                                                discId = 4
                                                                
                                                                End If
                                                                If discId <> 4 Then
                                         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = discId
                                          .TextMatrix(LngRow, .ColIndex("DiscountType")) = discId
                                          .TextMatrix(LngRow, .ColIndex("DiscountVal")) = discValue
                                          End If
                                                                                    
                                                                        
                                End If
                                            
                                            End If
                        
                        
                     
                                
                                
                                
                        End If
                        


         If Grid.ColIndex("ParrtNoCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ParrtNoCode")) = ParrtNoCode
                           End If
                      
                      
                      If Grid.ColIndex("ItemDetailedCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ItemDetailedCode")) = ItemDetailedCode
                           End If
                           
                           
            '-----------------------------------
            If SystemOptions.UserInterface = ArabicInterface Then
                .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
            Else
                .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
            End If

            '----------------------------------
            '--------??C?E C????C? C?????E EC????
            If RsItems("RelatedItem").value = True Then
                Set Frm = New FrmAddItemAttachs
                Frm.LoadItemData LngItemID, DblQty
                Frm.show vbModal

                If Frm.UserCancelled = True Then
                    Unload Frm
                Else

                    For j = Frm.FG.FixedRows To Frm.FG.Rows - 1

                        If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                            LngAttachedItemID = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                            DblAttachedQty = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                        If SystemOptions.Itemsattachedzero = False Then
                            DblAttachedPrice = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                         Else
                             DblAttachedPrice = 0
                         End If
                         
                            AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice, , , , , , , , , , , , , , 4
                            
                            
              



                        End If

                    Next j

                    Unload Frm
                End If
            End If

            '----------------------------------
            Me.Grid.ShowCell LngRow, Me.Grid.ColIndex("Name")
            ReSerialGrid
            Me.Calculate LngRow
           
            '====================================================================
  
            Set RsUnitData = New ADODB.Recordset
            StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName, TblUnites.UnitNamee," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
            StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
            StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
            StrSQL = StrSQL + " AND DefaultUnit=1"
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                If DefalutUnitID = 0 Then
                    .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").value
                    longunitid = RsUnitData("UnitID").value
                    
           If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").value
          Else
                   .TextMatrix(LngRow, .ColIndex("UnitID")) = IIf(IsNull(RsUnitData("UnitNamee").value), RsUnitData("UnitName").value, RsUnitData("UnitNamee").value)
          End If
                Else
                    .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = DefalutUnitName
                End If
        
            End If

            RsUnitData.Close
            Set RsUnitData = Nothing
            PutDefCount LngRow
       
       
            
           '*******************************************19032017*********
                  .TextMatrix(LngRow, .ColIndex("Remarks")) = UnitName
                  .TextMatrix(LngRow, .ColIndex("QtyFaqtors")) = DescUnitFact(LngItemID, longunitid)
                  
          If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & "  ﬂÊ‰ „‰"
         .TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, longunitid)
            Else
            '.TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & "Consisting of"
            '.TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, longunitid)
            End If
           
        
  .TextMatrix(LngRow, .ColIndex("ReminRequ")) = GetReustValue(val(Me.StoreName.BoundText), LngItemID)
            
            '*******************************************19032017*********
       
       
            '====================================================================
        If subcall = False Then Exit Function
        
        
'free items qty
  If Not IsNull(RsItems("FreeQty").value) = True Then
  
  If val(RsItems("FreeQty").value) > 0 Then
       AddItemInGrid 0, LngItemID, DblQty * RsItems("FreeQty").value / 100, SngItemPrice, , , , , , 4, False
       subcall = False
       
  End If
  
  End If
           If Me.GridTrans = internalorder Or Me.GridTrans = purchaserequest Or Me.GridTrans = salespricelist Or GridTrans = ShowPrice Or Me.GridTrans = MoveItems Then
                       If Grid.ColIndex("ItemBalance") <> -1 Then
           '                 GetMonthlyAverage 19, (LngItemID), Fromdate, Date, ItemBalance
           '                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = ItemBalance
              
  '  Dim RsTest As ADODB.Recordset
  '  Dim DblOutQty As Double
  '  Dim LngColorID As Long
  '  Dim LngClassId As Long
'
'    Dim StrItemSize As String
'                                   LngColorID = ColorID ' val(.TextMatrix(LngRow, .ColIndex("ColorID")))
'                                StrItemSize = itemsize ' .TextMatrix(LngRow, .ColIndex("ItemSize"))
'                                LngClassId = ClassId ' val(.TextMatrix(LngRow, .ColIndex("ClassId")))
'
'                                If Trim(Grid.TextMatrix(LngRow, .ColIndex("Serial"))) <> "" Then
'                                    StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))
'                                Else
'                                    StrItemSerial = ""
'                                End If
'                          '
'                        If Me.GridTrans = purchaserequest Then
'                           .TextMatrix(LngRow, .ColIndex("ItemSalesYear")) = GetQtySqlYearly(LngItemID, val(val(Me.StoreName.BoundText)), year(TransactionDate), LngColorID, StrItemSize, LngClassId)
'                        End If
'            Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), TransactionDate, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
'              If RsTest.EOF Or RsTest.BOF Then
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = 0
'              Else
'              ' (Round(RsTest("totalqty").value, 2)
'
'              ItemBalance = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
'
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = Round(ItemBalance, 2)
'              End If
'
           End If
          End If
        End With

    End If
    GetItemBalanceInStore LngRow, ColorID, itemsize, ClassId, StrItemSerial, LngItemID, TransactionDate
End Function
Public Function GetQty(Optional ItemID As Long, Optional CurRow As Long) As Double
Dim i As Long
Dim SumQty As Double
SumQty = 0
With Grid
For i = 1 To .Rows - 1
If ItemID = val(.TextMatrix(i, .ColIndex("Code"))) And i <> CurRow And .TextMatrix(i, .ColIndex("ExpiryDate")) <> "" And val(.TextMatrix(i, .ColIndex("IsExpirDate"))) <> 1 Then
SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
End If
Next i
End With
GetQty = SumQty
End Function

 Public Sub ItemLotNoFunFIFo(LngRow As Long, _
                         LngItemID As Long, _
                          Optional DblQty As Double = 1, _
                              Optional SngItemPrice As Single = 0, _
                              Optional IntItemCase As Integer = 0, _
                              Optional StrItemSerial As String = "", _
                              Optional FromHeader As Boolean = False, _
                              Optional DefalutUnitID As Long = 0, _
                              Optional DefalutUnitName As String, Optional freeitemsqty As Integer = 0, _
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer, Optional StoreId As Integer, Optional Qty2 As Double = 0)
    
        Dim QtyFiled As Double
        Dim i As Integer
        Dim rs As ADODB.Recordset
        Dim My_SQL As String
        Dim k As Integer
        Dim Qty As Double
        Dim TempQty As Double
        Dim lotNo As String
        Dim ExpiryDate As String
        Dim ProductionDate As String
        Qty = DblQty
        Set rs = New ADODB.Recordset
        My_SQL = " SELECT     TOP 100 PERCENT  dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) "
        My_SQL = My_SQL & " AS Qty ,dbo.Transaction_Details.ProductionDate "
        My_SQL = My_SQL & " FROM         dbo.Transaction_Details INNER JOIN"
        My_SQL = My_SQL & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        My_SQL = My_SQL & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
        My_SQL = My_SQL & " Where (dbo.TransactionTypes.StockEffect <> 0) "
        My_SQL = My_SQL & " and (dbo.Transactions.StoreID = " & StoreId & ")"
        My_SQL = My_SQL & " And (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
          If Me.GridTrans = InvoiceTransaction Then
            My_SQL = My_SQL + " and dbo.Transactions.Transaction_ID<>" & val(Me.TxtNots)
        Else
            My_SQL = My_SQL + " and dbo.Transactions.Transaction_ID<>" & val(Me.TxtInvID)
        End If
 
        My_SQL = My_SQL & " GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.ProductionDate"
       My_SQL = My_SQL + " Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) > 0)"
        My_SQL = My_SQL & " ORDER BY dbo.Transaction_Details.ExpiryDate"

        rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
     If rs.RecordCount > 0 Then
Dim FlgEqula As Boolean
With Grid
If val(.TextMatrix(LngRow, .ColIndex("IsExpirDate"))) = 1 Then Exit Sub
FlgEqula = True
 k = 1
                    For i = 1 To rs.RecordCount
                    TempQty = 0
                    If Qty2 <> 0 Then
                 
                    QtyFiled = IIf(IsNull(rs.Fields("Qty").value), 0, rs.Fields("Qty").value)
                    If Qty2 > QtyFiled Then
                    TempQty = QtyFiled
                    Qty2 = Qty2 - QtyFiled
                    Else
                    TempQty = QtyFiled - Qty2
                    Qty2 = 0
                    FlgEqula = False
                    End If
                    End If
    If TempQty = 0 And FlgEqula = False Then
    FlgEqula = True
    If i = rs.RecordCount Then
       ' If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
            .Rows = .Rows + 1
       ' End If
                        
                  '      AddItemInGrid2 .Rows - 1, LngItemID, TempQty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, StoreID
    End If
    GoTo X
    End If
                    If Qty2 = 0 Then
                    Qty2 = TempQty
                    If k = 1 Then
                    If Qty2 > 0 Then
                    QtyFiled = Qty2
                    Else
                    QtyFiled = IIf(IsNull(rs.Fields("Qty").value), 0, rs.Fields("Qty").value)
                    End If
                    Else
                    QtyFiled = IIf(IsNull(rs.Fields("Qty").value), 0, rs.Fields("Qty").value)
                    End If
                    If Qty <> 0 Then
                    If Qty > QtyFiled Then
                    TempQty = QtyFiled
                    Qty = Qty - QtyFiled
                    Else
                    TempQty = Qty
                    Qty = 0
                    End If
                    
                    lotNo = IIf(IsNull(rs.Fields("LotNO").value), "", rs.Fields("LotNO").value)
                    ExpiryDate = IIf(IsNull(rs.Fields("ExpiryDate").value), "", rs.Fields("ExpiryDate").value)
                    ProductionDate = IIf(IsNull(rs.Fields("ProductionDate").value), "", rs.Fields("ProductionDate").value)
                    If k <> 1 Then
                    If TempQty <> 0 Then
                    .Rows = .Rows + 1
                     AddItemInGrid2 .Rows - 1, LngItemID, TempQty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, StoreId, lotNo, ExpiryDate, ProductionDate
                     End If
                    Else
                    If TempQty <> 0 Then
                    k = 0
                        .TextMatrix(LngRow, .ColIndex("LotNO")) = lotNo
                        .TextMatrix(LngRow, .ColIndex("ExpiryDate")) = ExpiryDate
                        .TextMatrix(LngRow, .ColIndex("ProductionDate")) = ProductionDate
                        .TextMatrix(LngRow, .ColIndex("Count")) = TempQty
                      ReCalVAT LngRow
                    Else
                    .RemoveItem LngRow
                    End If
                    End If
                        
                      End If
                      End If
X:
                      rs.MoveNext

                    Next i
         If Qty > 1 And SngItemPrice > 0 Then
           'If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                           .Rows = .Rows + 1
                           
        'End If
               
                     AddItemInGrid2 .Rows - 1, LngItemID, Qty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, StoreId
                  ReCalVAT LngRow
         End If
End With
                    rs.Close
   End If
   
     ColorExperDate
     SentTypeVAT
End Sub
 
   

  Public Function AddItemInGrid(LngRow As Long, _
                              LngItemID As Long, _
                              Optional DblQty As Double = 1, _
                              Optional SngItemPrice As Single = 0, _
                              Optional IntItemCase As Integer = 0, _
                              Optional StrItemSerial As String = "", _
                              Optional FromHeader As Boolean = False, _
                              Optional DefalutUnitID As Long = 0, _
                              Optional DefalutUnitName As String, Optional freeitemsqty As Integer = 0, _
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer)

    Dim RsParts As ADODB.Recordset
    Dim StrSQL  As String
    Dim i As Long, j As Long
    Dim DblPartQty As Double
    Dim longunitid As Long
    Dim UnitName As String
    Dim LngPartItemIndex As Long
    Dim LngNewPartRow As Long
    Dim LngAddCount As Long
    Dim StrSerial As String
    Dim Frm As FrmAddItemAttachs
    Dim LngAttachedItemID As Long
    Dim DblAttachedQty As Double
    Dim DblAttachedPrice As Single
    Dim RsUnitData As ADODB.Recordset
    Dim LngUnitID As Long
                 '------------------  Â‰« «·«÷«›…
    If Me.GridTrans = NewGard Then
        Dim LngFindRow As Long
        LngFindRow = Grid.FindRow(LngItemID, 0, Grid.ColIndex("Code"), False, True)
        If LngFindRow <> -1 Then
            LngRow = LngFindRow
            DblQty = DblQty + val(Grid.TextMatrix(LngRow, Grid.ColIndex("Count")))
        End If
    
    End If
    '------------------  Â‰« «·«÷«›…
    
     
              
    If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = ProductionOrder) And SystemOptions.AllowItemByRow = True Then
                               
    If CheckMaultyItem(LngRow, LngItemID, DefalutUnitID, DblQty) = False Then
    If LngRow = 0 Then
        LngRow = Me.NewEmptyRow
    ElseIf LngRow > Me.Grid.Rows - 1 Then
        Me.Grid.Rows = LngRow + 1
    End If
    Else
    Me.Calculate LngRow, , , True
End If
Else
   If LngRow = 0 Then
        LngRow = Me.NewEmptyRow
    ElseIf LngRow > Me.Grid.Rows - 1 Then
        Me.Grid.Rows = LngRow + 1
    End If
End If
    If Not (RsItems.BOF Or RsItems.EOF) Then
        If RsItems("ItemID").value <> LngItemID Then
            RsItems.find "ItemID=" & LngItemID & "", , adSearchForward, 1
        End If
    End If

    If Not (RsItems.BOF Or RsItems.EOF) Then

        '«·’‰› «·„Ã„⁄
        With Me.Grid
  
            If RsItems("AssbliedItem").value = True Then
                '   StrSQL = "SELECT  unitid,TableID, ItemID, PartItemID, PartItemQty, PartItemPrice"
                '   StrSQL = StrSQL + " FROM TblItemsParts Where ItemID=" & LngItemID
                '   StrSQL = StrSQL + " Order By TableID"
            
                StrSQL = "SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
                StrSQL = StrSQL + " dbo.TblItemsParts.ItemID , dbo.TblItemsParts.TableID, dbo.TblUnites.unitname, dbo.TblUnites.UnitNamee"
                StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
                StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID"
                StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & LngItemID & ")"
                StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"
                Set RsParts = New ADODB.Recordset
                RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsParts.EOF Or RsParts.BOF) Then

                    For i = 0 To RsParts.RecordCount - 1
                        LngPartItemIndex = LngPartItemIndex + 1
                        '--------------------------------------
                        'Add to grid the New Parts Rows
                        '--------------------------------------
 
                        longunitid = IIf(IsNull(RsParts("UnitID").value), 0, RsParts("UnitID").value)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            UnitName = IIf(IsNull(RsParts("unitname").value), 0, RsParts("unitname").value)
                        Else
                            UnitName = IIf(IsNull(RsParts("unitnamee").value), 0, RsParts("unitnamee").value)
                        End If

                        DblPartQty = IIf(IsNull(RsParts("PartItemQty").value), 0, RsParts("PartItemQty").value)
                        DblPartQty = DblPartQty * DblQty
                        LngNewPartRow = LngRow + (LngPartItemIndex - 1)
                        Me.Grid.AddItem "", LngNewPartRow
                        RsItems.MoveFirst
                        AddItemInGrid LngNewPartRow, RsParts("PartItemID").value, DblPartQty, , , , , longunitid, UnitName
                        RsParts.MoveNext
                    Next i

                End If
            
                '--------≈÷«›… «·√’‰«› «·„·Õﬁ… »«·’‰›
                If Not (RsItems.BOF Or RsItems.EOF) Then
                    If RsItems("ItemID").value <> LngItemID Then
                        RsItems.find "ItemID=" & LngItemID & "", , adSearchForward, 1
                    End If
                End If

                If RsItems("RelatedItem").value = True Then
                    Set Frm = New FrmAddItemAttachs
                    Frm.LoadItemData LngItemID, DblQty
                    Frm.show vbModal

                    If Frm.UserCancelled = True Then
                        Unload Frm
                    Else

                        For j = Frm.FG.FixedRows To Frm.FG.Rows - 1

                            If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                                LngAttachedItemID = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                                DblAttachedQty = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                                DblAttachedPrice = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                        AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice, , , , , , , , , , , , , , 4
             
                         
                            
                            End If

                        Next j

                        Unload Frm
                    End If
                End If

                Exit Function
            End If

            .TextMatrix(LngRow, .ColIndex("Code")) = LngItemID
            .TextMatrix(LngRow, .ColIndex("Name")) = LngItemID
            .TextMatrix(LngRow, .ColIndex("HaveSerial")) = RsItems("HaveSerial").value
            If Me.GridTrans = PurchaseTransaction Or Me.GridTrans = Returntransaction Then
               If SystemOptions.LinkSupplerWithItem = True Then
                Customer.BoundText = GetSuppler(LngItemID)
               
               End If
            
            End If
            
                              
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Then
                If SystemOptions.IsGeometricProportions Then
                    LngUnitID = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
                   
                    If LngUnitID = 0 Then
                        GetDefaultItemUnit LngItemID, LngUnitID
                    End If
                     StrSQL = "SELECT     TOP 100 PERCENT PartItemQty,ForUnit,MethodCalc "
                    
                    StrSQL = StrSQL + " FROM         dbo.TblItemsUnits"
                    
                    StrSQL = StrSQL + " Where (dbo.TblItemsUnits.ItemID = " & LngItemID & ")"
                    StrSQL = StrSQL + " and (dbo.TblItemsUnits.UnitID = " & LngUnitID & ")"
                    
                    Set RsParts = New ADODB.Recordset
                    RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
                    If Not (RsParts.EOF Or RsParts.BOF) Then
                        .TextMatrix(LngRow, .ColIndex("ItemQty")) = RsParts!PartItemQty & ""
                        .TextMatrix(LngRow, .ColIndex("ForUnit")) = RsParts!ForUnit & ""
                        .TextMatrix(LngRow, .ColIndex("MethodCalc")) = RsParts!MethodCalc & ""
                    End If
                End If
                checkQtybylengthoption LngRow, LngItemID
            End If
 
            If Not Customer Is Nothing Then
'Wael Updated
                .TextMatrix(LngRow, .ColIndex("CusID2")) = Customer.BoundText
                If Me.GridTrans2 = InvoiceTransactionCompose Then
                    Customer.BoundText = ""
                   ' txtCustCode.SetFocus
                End If
            End If
If Not TxtLotNo Is Nothing Then
            .TextMatrix(LngRow, .ColIndex("LotNO")) = Me.TxtLotNo.Text
            Me.TxtLotNo.Text = ""
End If
'.TextMatrix(LngRow, .ColIndex("DiscountType")) = 4

                        
                        
                        
                           
    If .ColIndex("showPrice2") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("showPrice2")) = GetDefaultGoldPrice
            End If
            

            If .ColIndex("guaranteeTime") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("guaranteeTime")) = IIf(IsNull(RsItems("GuaranteeValue").value), 0, RsItems("GuaranteeValue").value)
            End If

            'DblQty = Val(.TextMatrix(LngRow, .ColIndex("Count")))
            DblQty = IIf(DblQty = 0, 1, DblQty)
            .TextMatrix(LngRow, .ColIndex("Count")) = DblQty
            .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1 'set the defalut item case to new case

  If Me.Grid.ColIndex("ItemType") <> -1 Then
   .TextMatrix(LngRow, .ColIndex("ItemType")) = RsItems("ItemType").value
  End If
            If .ColIndex("DiscountType") <> -1 Then
            
            If freeitemsqty = 0 Then
                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
             Else
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
             
         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                    .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                    
                    
             End If
            End If

      If .ColIndex("ExtraType") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraType")) = 1
            End If
             
            If .ColIndex("order_no") <> -1 Then
                ' .TextMatrix(LngRow, .ColIndex("order_no")) = GetItemOrderNo(LngItemID)
            End If
                   
                   GetDefaultItemUnit LngItemID, LngUnitID
                .TextMatrix(LngRow, .ColIndex("SalesValue")) = GetItemPrice(LngItemID, 1, LngUnitID)
            If SngItemPrice = 0 Then
                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Then
                    'Put Sales Price
                    
                    
                    
                .TextMatrix(LngRow, .ColIndex("Price")) = GetItemPrice(LngItemID, 1, LngUnitID)
               ' ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
               ElseIf GridTrans = ReturnDestruction Or Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
               
                    '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                    StrSerial = Trim$(.TextMatrix(LngRow, .ColIndex("Serial")))
                    Dim UnitID As Long
 
                    UnitID = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                    .TextMatrix(LngRow, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), UnitID, val(Me.StoreName.BoundText))
                
                
                ElseIf Me.GridTrans = ProductionOrder Then
                    .TextMatrix(LngRow, .ColIndex("Price")) = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
          
                
            ElseIf Me.GridTrans = internalorder Or Me.GridTrans = purchaserequest Then
                
                
               
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    Fromdate = FirstPeriodDateInthisYear
 
    
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Or Grid.ColIndex("LastPurchaseqty") <> -1 Or Grid.ColIndex("LastPurchasePrice") <> -1 Or Grid.ColIndex("RequestLimit") <> -1 Then
                           
                           GetlastPurchasedata 20, (LngItemID), Fromdate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                           End If
                           
                           
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchaseDate")) = LastPurchaseDate
                           End If
                           
                                If Grid.ColIndex("LastPurchasePrice") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchasePrice")) = LastPurchasePrice
                           End If
                           
                    If Grid.ColIndex("LastPurchaseqty") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("LastPurchaseqty")) = LastPurchaseqty
                           End If
                           
                               If Grid.ColIndex("RequestLimit") <> -1 Then
                               .TextMatrix(LngRow, .ColIndex("RequestLimit")) = IIf(IsNull(RsItems("RequestLimit").value), 0, RsItems("RequestLimit").value)
                           End If
                           
                           
                            If Grid.ColIndex("AverageIssue") <> -1 Then
                            GetMonthlyAverage 19, (LngItemID), Fromdate, Date, AverageMonthly
                               .TextMatrix(LngRow, .ColIndex("AverageIssue")) = AverageMonthly
                           End If
                           
                           
                                   If Grid.ColIndex("AverageIssueyraly") <> -1 Then
                            GetYearlyAverage 19, (LngItemID), Fromdate, Date, AverageIssueyraly
                               .TextMatrix(LngRow, .ColIndex("AverageIssueyraly")) = AverageIssueyraly
                           End If
                           
                           
      '                 If Grid.ColIndex("ItemBalance") <> -1 Then
           '                 GetMonthlyAverage 19, (LngItemID), Fromdate, Date, ItemBalance
           '                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = ItemBalance
              
  '  Dim RsTest As ADODB.Recordset
  '  Dim DblOutQty As Double
  '  Dim LngColorID As Long
  '  Dim LngClassId As Long
'
'    Dim StrItemSize As String
'                                   LngColorID = ColorID ' val(.TextMatrix(LngRow, .ColIndex("ColorID")))
'                                StrItemSize = itemsize ' .TextMatrix(LngRow, .ColIndex("ItemSize"))
'                                LngClassId = ClassId ' val(.TextMatrix(LngRow, .ColIndex("ClassId")))
'
'                                If Trim(Grid.TextMatrix(LngRow, .ColIndex("Serial"))) <> "" Then
'                                    StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))
'                                Else
'                                    StrItemSerial = ""
'                                End If
'                           .TextMatrix(LngRow, .ColIndex("ItemSalesYear")) = GetQtySqlYearly(LngItemID, val(val(Me.StoreName.BoundText)), year(TransactionDate), LngColorID, StrItemSize, LngClassId)
'            Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), TransactionDate, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
'              If RsTest.EOF Or RsTest.BOF Then
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = 0
'              Else
'              ' (Round(RsTest("totalqty").value, 2)
'
'              ItemBalance = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
'
'                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = Round(ItemBalance, 2)
'              End If
'
'
'                           End If
'

                Else
                    'Put Purchase Price
                    .TextMatrix(LngRow, .ColIndex("Price")) = IIf(IsNull(RsItems("PurchasePrice").value), 0, RsItems("PurchasePrice").value)
                End If

            Else
                'Put the Sended Price
                .TextMatrix(LngRow, .ColIndex("Price")) = SngItemPrice
            End If
            .TextMatrix(LngRow, .ColIndex("UnitPrice")) = .TextMatrix(LngRow, .ColIndex("Price"))
            If SystemOptions.IsGeometricProportions = True Then
                .TextMatrix(LngRow, .ColIndex("Price")) = (.TextMatrix(LngRow, .ColIndex("UnitPrice"))) / val(.TextMatrix(LngRow, .ColIndex("Count")))
            End If

            '----------------------------------
            If IntItemCase = 0 Then
                If IsNull(RsItems("ItemCase").value) Then
                    .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1
                Else
                    .TextMatrix(LngRow, .ColIndex("ItemCase")) = RsItems("ItemCase").value
                End If

            Else
                .TextMatrix(LngRow, .ColIndex("ItemCase")) = IntItemCase
            End If

            .TextMatrix(LngRow, .ColIndex("Serial")) = StrItemSerial
            '        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            '            .TextMatrix(LngRow, .ColIndex("Serial")) = TxtSerial.text
            '        End If
        
            If Me.GridTrans <> OpeningBalance And Me.GridTrans <> Returntransaction And Me.GridTrans <> NewGard Then
           '     .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
                
                
                
                      If freeitemsqty = 0 Then
                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
                .TextMatrix(LngRow, .ColIndex("DiscountVal")) = 0
             Else
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
       '      .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
             
         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                    .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                    
                    
             End If
                
                If Me.Grid.ColIndex("ExtraType") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraType")) = 1
                End If
                
                If Me.Grid.ColIndex("ExtraVal") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ExtraVal")) = 0
                End If
                If Me.Grid.ColIndex("ScurrencyID") <> -1 Then
                .TextMatrix(LngRow, .ColIndex("ScurrencyID")) = 1
                End If
                
            End If

            If Me.Grid.ColIndex("ItemType") <> -1 Then
                If Not (IsNull(RsItems("ItemType").value)) Then
                    .TextMatrix(LngRow, .ColIndex("ItemType")) = RsItems("ItemType").value

                    If val(.TextMatrix(LngRow, .ColIndex("ItemType"))) = 1 Then
                                            .Cell(flexcpPicture, LngRow, .ColIndex("Ser")) = _
                                                 mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
                    Else

                        If .ColIndex("Ser") <> -1 Then
                            .Cell(flexcpPicture, LngRow, .ColIndex("Ser")) = Nothing
                        End If
                    End If

                Else
                    .TextMatrix(LngRow, .ColIndex("ItemType")) = 0
                End If
            End If

            '-----------------------------------
            .TextMatrix(LngRow, .ColIndex("ColorID")) = ColorID
            .TextMatrix(LngRow, .ColIndex("ItemSize")) = itemsize
            .TextMatrix(LngRow, .ColIndex("Classid")) = ClassId
     '       .TextMatrix(LngRow, .ColIndex("StoreID")) = 1
     FillStoredLin
     
 .TextMatrix(LngRow, .ColIndex("LotNO")) = ""
             .TextMatrix(LngRow, .ColIndex("ExpiryDate")) = ""
             .TextMatrix(LngRow, .ColIndex("ProductionDate")) = ""
            .TextMatrix(LngRow, .ColIndex("IsExpirDate")) = ""

            
                          '«·’‰› «·„·Õﬁ €«·»« „Ã«‰Ì
                        If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) And discountype = 4 Then
                                If .ColIndex("DiscountType") <> -1 Then
                                        If SystemOptions.attacheditemsisfree = True Then
                                                                 .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                                                                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                                                   
                                        End If
                                End If
                                
                        End If
                        
'⁄—Ê÷ «·«’‰«›
Dim discId As Double
Dim discValue As Double
                        If (Me.GridTrans = InvoiceTransaction) Then
                                            If CheckItem(LngItemID, DtpBillDate, discId, discValue, , val(Branch.BoundText)) = True Then
 
                                                                      If .ColIndex("DiscountType") <> -1 Then
                                                                      
                                                                   '0 value 1 percentage
                                                                If discId = 0 Then
                                                                discId = 2
                                                                ElseIf discId = 1 Then
                                                                discId = 3
                                                                             ElseIf discId = 4 Then
                                                                discId = 4
                                                                
                                                                End If
                                                                If discId <> 4 Then
                                         .Cell(flexcpData, LngRow, .ColIndex("DiscountType")) = discId
                                          .TextMatrix(LngRow, .ColIndex("DiscountType")) = discId
                                          .TextMatrix(LngRow, .ColIndex("DiscountVal")) = discValue
                                          End If
                                                                                    
                                                                        
                                End If
                                            
                                            End If
                        
                        
                     
                                
                                
                                
                        End If
                        


         If Grid.ColIndex("ParrtNoCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ParrtNoCode")) = ParrtNoCode
                           End If
                      
                      
                      If Grid.ColIndex("ItemDetailedCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ItemDetailedCode")) = ItemDetailedCode
                           End If
                           
                           
            '-----------------------------------
            If SystemOptions.UserInterface = ArabicInterface Then
                .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
            Else
                .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
            End If

            '----------------------------------
            '--------≈÷«›… «·√’‰«› «·„·Õﬁ… »«·’‰›
            If RsItems("RelatedItem").value = True Then
                Set Frm = New FrmAddItemAttachs
                Frm.LoadItemData LngItemID, DblQty
                Frm.show vbModal

                If Frm.UserCancelled = True Then
                    Unload Frm
                Else

                    For j = Frm.FG.FixedRows To Frm.FG.Rows - 1

                        If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                            LngAttachedItemID = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                            DblAttachedQty = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                        If SystemOptions.Itemsattachedzero = False Then
                            DblAttachedPrice = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                         Else
                             DblAttachedPrice = 0
                         End If
                         
                            AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice, , , , , , , , , , , , , , 4
                            
                            
              



                        End If

                    Next j

                    Unload Frm
                End If
            End If

            '----------------------------------
            Me.Grid.ShowCell LngRow, Me.Grid.ColIndex("Name")
            ReSerialGrid
            Me.Calculate LngRow
           
            '====================================================================
  
            Set RsUnitData = New ADODB.Recordset
            StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName, TblUnites.UnitNamee," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
            StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
            StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
            StrSQL = StrSQL + " AND DefaultUnit=1"
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                If DefalutUnitID = 0 Then
                    .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").value
                    longunitid = RsUnitData("UnitID").value
                    
           If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").value
          Else
                   .TextMatrix(LngRow, .ColIndex("UnitID")) = IIf(IsNull(RsUnitData("UnitNamee").value), RsUnitData("UnitName").value, RsUnitData("UnitNamee").value)
          End If
                Else
                    .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = DefalutUnitName
                End If
        
            End If

            RsUnitData.Close
            Set RsUnitData = Nothing
            PutDefCount LngRow
       
       
            
           '*******************************************19032017*********
                  .TextMatrix(LngRow, .ColIndex("Remarks")) = UnitName
                  .TextMatrix(LngRow, .ColIndex("QtyFaqtors")) = DescUnitFact(LngItemID, longunitid)
                  
          If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & "  ﬂÊ‰ „‰"
         .TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, longunitid)
            Else
            '.TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & "Consisting of"
          '  .TextMatrix(LngRow, .ColIndex("Remarks")) = .TextMatrix(LngRow, .ColIndex("Remarks")) & " " & DescUnit(LngItemID, longunitid)
            End If
           
        
  .TextMatrix(LngRow, .ColIndex("ReminRequ")) = GetReustValue(val(Me.StoreName.BoundText), LngItemID)
            
            '*******************************************19032017*********
       
       
            '====================================================================
        If subcall = False Then Exit Function
        
        
'free items qty
  If Not IsNull(RsItems("FreeQty").value) = True Then
  
  If val(RsItems("FreeQty").value) > 0 Then
       AddItemInGrid 0, LngItemID, DblQty * RsItems("FreeQty").value / 100, SngItemPrice, , , , , , 4, False
       subcall = False
       
  End If
  
  End If
        

   
        End With



    End If



'*************************************
   If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = MoveItems Then
            If SystemOptions.AllowExperDateFIFO = True Then
             ItemLotNoFunFIFo LngRow, LngItemID, DblQty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, val(StoreName.BoundText), GetQty(LngItemID, LngRow)
            End If
   End If
   
   '*************************************
      GetItemBalanceInStore LngRow, ColorID, itemsize, ClassId, StrItemSerial, LngItemID, TransactionDate
      FillStoredLin
      CalculteMaxQtyUnit LngRow
      ReCalVAT LngRow
   '*************************************
End Function





Sub GetItemBalanceInStore(Optional LngRow As Long, Optional ColorID As Integer, Optional itemsize As Integer, Optional ClassId As Integer, Optional StrItemSerial As String, Optional LngItemID As Long, Optional TransactionDate As Variant)
   If Me.GridTrans = internalorder Or Me.GridTrans = purchaserequest Or Me.GridTrans = salespricelist Or GridTrans = ShowPrice Or Me.GridTrans = MoveItems Then
    Dim RsTest As ADODB.Recordset
    Dim DblOutQty As Double
    Dim LngColorID As Long
    Dim LngClassId As Long
    Dim StrItemSize As String
        LngColorID = ColorID
        StrItemSize = itemsize
         LngClassId = ClassId
With Grid
     If Grid.ColIndex("ItemBalance") <> -1 Then
                                If Trim(.TextMatrix(LngRow, .ColIndex("Serial"))) <> "" Then
                                    StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))
                                Else
                                    StrItemSerial = ""
                                End If
                          '
                        If Me.GridTrans = purchaserequest Then
                           .TextMatrix(LngRow, .ColIndex("ItemSalesYear")) = GetQtySqlYearly(LngItemID, val(val(Me.StoreName.BoundText)), year(TransactionDate), LngColorID, StrItemSize, LngClassId)
                        End If
            Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), TransactionDate, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
              If RsTest.EOF Or RsTest.BOF Then
                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = 0
              Else
              ' (Round(RsTest("totalqty").value, 2)
              
              ItemBalance = IIf(IsNull(RsTest("totalqty").value), 0, RsTest("totalqty").value)
              
                    .TextMatrix(LngRow, .ColIndex("ItemBalance")) = Round(ItemBalance, 2)
              End If
              
           End If
         
 End With
  End If
End Sub



 
 
Private Sub ReSerialGrid()
    Dim IntGrdLoopRows As Integer
    Dim IntSortCounter As Integer
    Dim IntSortedColIndex As Integer
    IntSortedColIndex = Me.Grid.ColIndex("Ser")

    If IntSortedColIndex <> -1 Then
        IntSortCounter = 1

        For IntGrdLoopRows = Me.Grid.FixedRows To Me.Grid.Rows - 1
            Grid.TextMatrix(IntGrdLoopRows, IntSortedColIndex) = IntSortCounter
            IntSortCounter = IntSortCounter + 1
        Next IntGrdLoopRows

        Grid.AutoSize IntSortedColIndex, IntSortedColIndex, False
    End If

    Me.CountItems
End Sub

Public Sub ShowInvProfDialog()
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngTransID As Long

    If Me.GridTrans <> InvoiceTransaction Then Exit Sub
 '   Load FrmInvProfit
 '   FrmInvProfit.CostMethod = SystemOptions.SysMainStockCostMethod

    StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "TransactionTypes.TransactionTypeName "
    StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN Transactions ON " & "TransactionTypes.Transaction_Type = Transactions.Transaction_Type "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    'FrmInvProfit.FgItems.Rows = Me.Grid.Rows
    'FrmInvProfit.FgItems.FixedRows = Me.Grid.FixedRows

    For i = Me.Grid.FixedRows To Me.Grid.Rows - 1
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemID")) = Me.Grid.Cell(flexcpText, i, Me.Grid.ColIndex("Code"))
    
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemCode")) = Me.Grid.Cell(flexcpTextDisplay, i, Me.Grid.ColIndex("Code"))
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemName")) = Me.Grid.Cell(flexcpTextDisplay, i, Me.Grid.ColIndex("Name"))
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("CostPrice")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("ItemCostPrice"))
'
        If SystemOptions.SysMainStockCostMethod = LastPurPriceType Then
            LngTransID = val(Me.Grid.TextMatrix(i, Me.Grid.ColIndex("PofTransID")))

            If LngTransID <> 0 Then
                If Not (rs.BOF Or rs.EOF) Then
                    rs.find "Transaction_ID=" & LngTransID & "", , adSearchForward, 1

                    If Not (rs.BOF Or rs.EOF) Then
'                        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("TransSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
'                        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("TransType")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                    Else
                        rs.MoveFirst
                    End If
                End If
            End If
        End If

'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("InvQty")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("Count"))
'
'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("InvQtyCost")) = val(FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("InvQty"))) * val(FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("CostPrice")))
'
'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemSalePrice")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("Price"))
'        'DiscountVal
'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("Discount")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("DiscountVal"))
'
'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemProfit")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("ItemProfit"))
'
'        FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("CalculateType")) = ""
    
    Next i

'    With FrmInvProfit.FgItems
'        .Cell(flexcpBackColor, .FixedRows, .ColIndex("TransType"), .Rows - 1, .ColIndex("TransType")) = &H808000
'        .Cell(flexcpBackColor, .FixedRows, .ColIndex("TransSerial"), .Rows - 1, .ColIndex("TransSerial")) = &H808000
'        .AutoSize 0, .Cols - 1, False
'        FrmInvProfit.lblInvProf.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemProfit"), .Rows - 1, .ColIndex("ItemProfit"))
'    End With

   ' FrmInvProfit.show vbModeless, mdifrmmain
    'ItemCode
    'ItemName
    'PurPrice
    'TransSerial
    'InvQty
    'InvQtyCost
    'ItemSalePrice
    'ItemProfit
    'CalculateType
End Sub

Public Sub CountItems()

    If Not LblItemsCount Is Nothing Then
        LblItemsCount.ToolTipText = "⁄œœ «·√’‰«› «·„ÊÃÊœ… ›Ï «·›« Ê—…"

        If Grid.Rows > 1 Then
            LblItemsCount.Caption = Grid.Aggregate(flexSTCount, Grid.FixedRows, Grid.ColIndex("Code"), Grid.Rows - 1, Grid.ColIndex("Code"))
        Else
            LblItemsCount.Caption = 0
        End If
    End If

End Sub

Public Function IsReaptedSerials() As Boolean
    Dim i As Long
    Dim LngFindRow As Long
    Dim StrSerial As String
    Dim LngSearchRow As Long

    With Me.Grid
        .Col = .ColIndex("Serial")
        .Sort = flexSortGenericAscending
        LngSearchRow = .FixedRows

        Do While LngSearchRow < (.Rows - 1)
            StrSerial = .TextMatrix(LngSearchRow, .ColIndex("Serial"))
            LngFindRow = .FindRow(StrSerial, LngSearchRow + 1, .ColIndex("Serial"), False, True)
        
            If LngFindRow <> -1 Then
                IsReaptedSerials = True
                Exit Function
            Else
                LngSearchRow = LngSearchRow + 1
            End If

        Loop

    End With

End Function

Private Sub ShowItemTip(ToolTipType As ItemRowToolTip, _
                        StrItemID As String, _
                        LngShowRow As Long, _
                        LngShowCol As Long, _
                        Optional bSetFocus As Boolean = False, _
                        Optional BolShowInGrid As Boolean = True, _
                        Optional LngHwnd As Long)
    
    '------------------------------

    '------------------------------

    Dim uPoint As POINTAPI
    Static lNoteRow As Long, lNoteCol As Long, r As Long, c As Long
    Dim lLeft  As Single
    Dim LTop As Single
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Acmd As ADODB.Command
    Dim XParItemID As ADODB.Parameter
    Dim XParTransType As ADODB.Parameter
    Dim i As Integer
    Dim LngItemID As Long
    Dim ItemTransInfo As LastItemTransInfo
    Dim xRect As RECT

    HideComment
   ' LngItemID = val(StrItemID)

    With Me.Grid

        If BolShowInGrid = True Then
            If LngShowRow <= -1 Then Exit Sub
            If LngShowCol <= -1 Then Exit Sub
            If lNoteRow = LngShowRow And lNoteCol = LngShowCol Then
                'If bSetFocus = False Then Exit Sub
            End If

            ClientToScreen Grid.hwnd, uPoint
            lLeft = (uPoint.X * Screen.TwipsPerPixelX) + .ColPos(LngShowCol) + .ColWidth(LngShowCol)
            LTop = (uPoint.Y * Screen.TwipsPerPixelY) + .RowPos(LngShowRow) + .RowHeight(LngShowRow) + 50
            'lLeft = lLeft + 100
        Else
            ClientToScreen LngHwnd, uPoint
            GetWindowRect LngHwnd, xRect
            lLeft = (uPoint.X * Screen.TwipsPerPixelX) ' + xRect.right '+ .ColPos(LngShowCol) + .ColWidth(LngShowCol)
            LTop = (uPoint.Y * Screen.TwipsPerPixelY) + xRect.bottom   '+ .RowPos(LngShowRow) + .RowHeight(LngShowRow) + 50
        End If
    
        Set FrmComment = New FrmFlexQtyTip

        If ToolTipType = ItemQtyToolTip Then

            With FrmComment
                .PicContainer(0).Visible = True
                .PicContainer(1).Visible = False
                .lbldata(0).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Item_Code"))
                .lbldata(1).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Item_Name"))
                .lbldata(2).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Group_Name"))
                '.LblData(3).Caption = Me.DcboStores.Text
                '--------------------------------
                'Get The Item Quantity
                '            Set Rs = GetItemQuantityStock(StrItemID, Me.DcboStores.BoundText, Date, False, Me.DcboStores.Text)
                '            If Not Rs Is Nothing Then
                '                If Not (Rs.BOF Or Rs.EOF) Then
                '                    Rs.MoveFirst
                '                    For I = 1 To Rs.RecordCount
                '                        Load FrmComment.LblUnit(I)
                '                        With FrmComment.LblUnit(I)
                '                            .Visible = True
                '                            .left = FrmComment.LblUnit(0).left
                '                            .top = FrmComment.LblUnit(I - 1).top + 30 + FrmComment.LblUnit(I - 1).Height
                '                            If SystemOptions.UserInterface = ArabicInterface Then
                '                                .Caption = IIf(IsNull(Rs("Unite_ArabicName").Value), "", Rs("Unite_ArabicName").Value)
                '                            Else
                '                                .Caption = IIf(IsNull(Rs("Unite_EnglishName").Value), "", Rs("Unite_EnglishName").Value)
                '                            End If
                '                            .ZOrder 0
                '                        End With
                '                        Load FrmComment.LblQty(I)
                '                        With FrmComment.LblQty(I)
                '                            .Visible = True
                '                            .left = FrmComment.LblQty(0).left
                '                            .top = FrmComment.LblUnit(I - 1).top + 30 + FrmComment.LblQty(I - 1).Height
                '                            .Caption = IIf(IsNull(Rs("Qty").Value), "", Rs("Qty").Value)
                '                            .ZOrder 0
                '                        End With
                '                        Rs.MoveNext
                '                    Next I
                '                End If
                '                Rs.Close
                '                Set Rs = Nothing
                '            End If
                '--------------------------------
                .left = lLeft
                .top = LTop
                ShowWindow .hwnd, SW_SHOWNA
                m_QtyToolTipShown = True
            End With

        Else

            With FrmComment
                .PicContainer(0).Visible = False
                .PicContainer(1).Visible = True
                'LngItemID = Val(Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Code")))
'                LngItemID = val(StrItemID)
                StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName," & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                StrSQL = StrSQL + " Where TblItems.ItemID='" & StrItemID & "'"
                
                Set rs = New ADODB.Recordset
                 rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .lbldata(7).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    .lbldata(6).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .lbldata(5).Caption = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                    .lbldata(4).Caption = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                    .lbldata(8).Caption = IIf(IsNull(rs("CustomerPrice").value), "", rs("CustomerPrice").value)
                    .lbldata(13).Caption = IIf(IsNull(rs("DealerPrice").value), "", rs("DealerPrice").value)
                    .lbldata(14).Caption = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
                End If

                If Me.GridTrans = InvoiceTransaction Then
                    ItemTransInfo = GetLastItemTrans(CLng(StrItemID), 21, val(val(Me.StoreName.BoundText)))
                    .lbl(4).Caption = "√Œ— ›« Ê—… »Ì⁄ ··’‰›:"
                    .lbl(10).Caption = "«”„ «·⁄„Ì·:-"
                ElseIf Me.GridTrans = PurchaseTransaction Then
                    ItemTransInfo = GetLastItemTrans(CLng(StrItemID), 22, val(val(Me.StoreName.BoundText)))
                    .lbl(4).Caption = "√Œ— ›« Ê—… ‘—«¡ ··’‰›:"
                    .lbl(10).Caption = "«”„ «·„Ê—œ:-"
                End If

                .lbldata(10).Caption = ItemTransInfo.TransactionSerial
                .lbldata(9).Caption = ItemTransInfo.TransactionDate
                .lbldata(11).Caption = ItemTransInfo.SngItemPrice
                .lbldata(12).Caption = ItemTransInfo.StrCustomerName
            
                .left = lLeft
                .top = LTop
                rs.Close
                Set rs = Nothing
                Set Acmd = Nothing
                Set XParItemID = Nothing
                Set XParTransType = Nothing
            
                ShowWindow .hwnd, SW_SHOWNA
                m_QtyToolTipShown = True
            End With

        End If

        m_QtyToolTipShown = True
        lNoteRow = LngShowRow
        lNoteCol = LngShowCol
    End With

End Sub

Private Sub HideComment()

    If Not FrmComment Is Nothing Then
        Unload FrmComment
        Set FrmComment = Nothing
        m_QtyToolTipShown = False
    End If

End Sub

Public Function CheckAvilableStock() As Boolean

End Function

Private Sub SetToolBar()

    If Not Me.GrdTBar Is Nothing Then

        With Me.GrdTBar
            .Buttons.Clear
            .Appearance = ccFlat
            .AllowCustomize = False
            .BorderStyle = ccNone
            .Style = tbrFlat
            .TextAlignment = tbrTextAlignBottom
            Set .ImageList = mdifrmmain.ImgLstTree
        
            .Buttons.Add , "AddRow", "", , "Plus"
            .Buttons.Add , "RemoveRow", , , "Minus"
            .Buttons.Add , "ClearAll", , , "Clear"
            .Buttons.Add , "FillData", , , "FillData"
            .Buttons.Add , "Tree", , MSComctlLib.ButtonStyleConstants.tbrCheck, "Tree"
            '.Buttons.Add , , , tbrSeparator
            '.Buttons.Add , "Options", , , "Design"
            .Buttons.Add , , , tbrSeparator
            .Buttons.Add , "Barcode", , tbrDefault, "barcode"
            .Buttons.Add , , , tbrSeparator
            .Buttons.Add , "GridCols", , tbrDefault, "GridOptions"
        
            If SystemOptions.UserInterface = ArabicInterface Then
                .Buttons("AddRow").ToolTipText = "≈÷«›… ”ÿ— ÃœÌœ ›Ï ««·›« Ê—…"
                .Buttons("RemoveRow").ToolTipText = "Õ–› «·”ÿ— «·„Õœœ «Ê «·Õ«·Ï ›Ï «·›« Ê—…"
                .Buttons("ClearAll").ToolTipText = "Õ–› ﬂ· «·”ÿÊ— «·„ÊÃÊœ… ›Ï «·›« Ê—…"
                .Buttons("FillData").ToolTipText = "≈œŒ«· «·√’‰«› ⁄‰ ÿ—Ìﬁ „—»⁄ ÕÊ«—"
                .Buttons("Tree").ToolTipText = "≈œŒ«· «·√’‰«› ⁄‰ ÿ—Ìﬁ ‘Ã—… «·√’‰«›"
                '.Buttons("Options").ToolTipText = "≈ŸÂ«— ŒÌ«—«  «·⁄—÷ ›Ï ÃœÊ· «·√’‰«›"
                .Buttons("Barcode").ToolTipText = "ÿ»«⁄… »«—ﬂÊœ ·√’‰«› «·›« Ê—…"
                .Buttons("GridCols").ToolTipText = "≈ŸÂ«— Ê≈Œ›«¡ «·√⁄„œ… ›Ï ÃœÊ· «·√’‰«›"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .Buttons("AddRow").ToolTipText = "Add new row in grid"
                .Buttons("RemoveRow").ToolTipText = "Remove the current select grid row"
                .Buttons("ClearAll").ToolTipText = "Remove all rows From the grid"
                .Buttons("FillData").ToolTipText = "Enter items in the grid with use the dialog box"
                '.Buttons("Options").ToolTipText = "Show & Hide Columns"
                .Buttons("Tree").ToolTipText = "Enter items in the grid with use the items tree"
                .Buttons("Barcode").ToolTipText = "Print Barcode Label..."
                .Buttons("GridCols").ToolTipText = "Show And Hide Grid Cols"
            End If

        End With

    End If

End Sub

Public Function NewEmptyRow() As Long
    Dim i As Long
    'Dim LngNewEmptyRow As Long

    For i = Me.Grid.FixedRows To Me.Grid.Rows - 1

        If val(Me.Grid.TextMatrix(i, Grid.ColIndex("Name"))) = 0 Then
            NewEmptyRow = i
            Exit Function
        End If

    Next i

    Me.Grid.Rows = Me.Grid.Rows + 1
    NewEmptyRow = Me.Grid.Rows - 1

End Function

Private Sub CloseMyAddItemForm()

    If Not MyAddItemForm Is Nothing Then
        Unload MyAddItemForm
        Set MyAddItemForm = Nothing
    End If

End Sub

Public Function ShowNewItemsPrice() As Boolean
    Dim Msg  As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim LngItemID As Long
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim LngFindRow As Long

    If Not m_FrmItemsNewPrices Is Nothing Then
        ShowNewItemsPrice = False
        Exit Function
    End If

    Set m_FrmItemsNewPrices = New FrmItemsNewPrices
    Set FG = m_FrmItemsNewPrices.FG
    FG.Rows = FG.FixedRows

    Set rs = New ADODB.Recordset
    rs.Open "TblItems", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If Not rs.BOF Or rs.EOF Then
        rs.MoveFirst

        For i = 1 To Me.Grid.Rows - 1
            LngItemID = val(Me.Grid.TextMatrix(i, Me.Grid.ColIndex("Name")))

            If LngItemID <> 0 Then
                LngFindRow = FG.FindRow(LngItemID, FG.FixedRows, FG.ColIndex("ItemID"), False, True)

                If LngFindRow = -1 Then
                    rs.MoveFirst
                    rs.find "ItemID=" & LngItemID & "", , adSearchForward, 1

                    If Not (rs.BOF Or rs.EOF) Then

                        With FG
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, .ColIndex("Serial")) = .Rows - 1
                            .TextMatrix(.Rows - 1, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                            .TextMatrix(.Rows - 1, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                            .TextMatrix(.Rows - 1, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                            .TextMatrix(.Rows - 1, .ColIndex("SallingPrice")) = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                            .TextMatrix(.Rows - 1, .ColIndex("CustomerPrice")) = IIf(IsNull(rs("CustomerPrice").value), "", rs("CustomerPrice").value)
                            .TextMatrix(.Rows - 1, .ColIndex("DealerPrice")) = IIf(IsNull(rs("DealerPrice").value), "", rs("DealerPrice").value)
                        End With

                    End If
                End If
            End If

        Next i

    End If

    FG.AutoSize 0, FG.Cols - 1, False
    m_FrmItemsNewPrices.show
End Function

Private Sub CloseNewItemsPrices()

    If Not m_FrmItemsNewPrices Is Nothing Then
        Unload m_FrmItemsNewPrices
        Set m_FrmItemsNewPrices = Nothing
    End If

End Sub

 
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

            If val(.TextMatrix(LngRowNum, .ColIndex("Code"))) <> 0 Then
                StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
  LngUnitID = val(Grid.Cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")))
                If LngUnitID = 0 Then
                 GetDefaultItemUnit LngItemID, LngUnitID
                 Grid.TextMatrix(LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 Grid.Cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 
                End If
               
'---------------«· «ﬂœ „‰ «‰ „œ… «·⁄—÷ ﬁÌ„… —ﬁ„Ì…
If Grid.ColIndex("PODaysSub") <> -1 Then
             If Not IsNumeric(val((Grid.TextMatrix(LngRowNum, Grid.ColIndex("PODaysSub"))))) Then
                                      Msg = " «—ÌŒ «· ”·Ì„  " & CHR(13)
                                      Msg = Msg + StrCurrentItemName & CHR(13)
                                      Msg = Msg + "€Ì— —ﬁ„Ì…"
                                      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                      FocusGridOnCell LngRowNum, Grid.ColIndex("PODaysSub")
                                      Screen.MousePointer = vbDefault
                                      Exit Function
            End If
End If
                '----------«· «ﬂœ „‰ ≈œŒ«· ﬂ„Ì… «·’‰›
                If val(.TextMatrix(LngRowNum, .ColIndex("Count"))) <= 0 And Me.GridTrans <> NewGard Then
                    Msg = "„‰ ›÷·ﬂ ÌÃ»  ÕœÌœ «·ﬂ„Ì… «·Œ«’… »«·’‰›" & CHR(13)
                    Msg = Msg + StrCurrentItemName & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    FocusGridOnCell LngRowNum, Grid.ColIndex("Count")
                    Screen.MousePointer = vbDefault
                    Exit Function
                    '-------«· «ﬂœ „‰  ≈œŒ«· ”⁄— «·’‰›
                    '
                ElseIf val(.TextMatrix(LngRowNum, .ColIndex("Price"))) <= 0 And (Me.GridTrans <> INVENTORYIN And Me.GridTrans <> ProductionOrder And Me.GridTrans <> ProductionMaterialReciveVoucher And Me.GridTrans <> InventoryOut And Me.GridTrans <> OpeningBalance And Me.GridTrans <> purchaseorderrequest And Me.GridTrans <> NewGard And Me.GridTrans <> internalorder And Me.GridTrans <> BookInventories And Me.GridTrans <> MoveItems And Me.GridTrans <> salespricelistRequest And Me.GridTrans <> ShipmentOrder And GridTrans <> ShipmentRegisteration And Me.GridTrans <> ShipmentRecieveVoucher And Me.GridTrans <> RecervieProductionVoucherNew And Me.GridTrans <> 24 And Me.GridTrans <> purchaserequest) Then
                    'Me.GridTrans <> PurchaseTransaction And
 'And Me.GridTrans <> Inventoryin
 
                    If Grid.ColIndex("DiscountType") <> -1 Then
                        
                        If val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) <> 4 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Price For Item" & CHR(13)
                        End If
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Price")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Price")
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If

                    Else
           '             Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
                      If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Price For Item" & CHR(13)
                        End If
                        
                        Msg = Msg + StrCurrentItemName & CHR(13)
                        Grid.Row = LngRowNum
                        Grid.Col = Grid.ColIndex("Price")
                        Grid.ShowCell LngRowNum, Grid.ColIndex("Price")
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Grid.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If

                    '-------«· «ﬂœ „‰ ≈œŒ«· ‰Ê⁄ «·Œ’„ ⁄·Ï «·»—‰«„Ã
                ElseIf Grid.ColIndex("DiscountType") <> -1 Then

                    If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                          Msg = "„‰ ›÷·ﬂ Õœœ ‰Ê⁄ «·Œ’„ «·Œ«’ »«·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Discount Type For Item" & CHR(13)
                        End If
                        
                        
                        Msg = Msg + StrCurrentItemName & CHR(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Grid.Row = LngRowNum
                        Grid.Col = Grid.ColIndex("DiscountType")
                        Grid.ShowCell LngRowNum, Grid.ColIndex("DiscountType")
                        Grid.SetFocus
                        Exit Function
                        '-------«· «ﬂœ „‰ ≈œŒ«· ﬁÌ„… «·Œ’„ „Ÿ»Êÿ
                    ElseIf (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 2 Or (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 3 Then

                        If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal"))) = "" Then
                       
                                             If SystemOptions.UserInterface = ArabicInterface Then
                           Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Discount Value For Item" & CHR(13)
                        End If
                        
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("DiscountVal")
                            Screen.MousePointer = vbDefault
                            Exit Function
                        ElseIf Not IsNumeric((Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal")))) Then
                        
                       '     Msg = "ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                                             If SystemOptions.UserInterface = ArabicInterface Then
                           Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Discount Value For Item" & CHR(13)
                        End If
                       
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = Msg + "€Ì— —ﬁ„Ì…"
                        Else
                        Msg = "Not Numeric" & CHR(13)
                        End If
                       
                       
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                            Msg = "«·’‰› " & CHR(13)
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            Msg = Msg + "·Â ”Ì—Ì«·" & CHR(13)
                            Msg = Msg + "„‰ ›÷·ﬂ √œŒ· «·”Ì—Ì«· «·Œ«’ »Â"
                        Else
                        
                        Msg = "Item " & CHR(13)
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            Msg = Msg + "Have Serial" & CHR(13)
                            Msg = Msg + "Please Enter Serial"
                        End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("Serial")
                            Screen.MousePointer = vbDefault
                            Exit Function
                            '›Ï ‰Ÿ«„ «·”Ì—Ì«· ·«»œ «‰ ÌﬂÊ‰ ﬂ· —ﬁ„ ›Ï Œ«‰… „‰›’·…
                        ElseIf val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Count"))) > 1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·’‰› " & CHR(13)
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            Msg = Msg + "·Â ”Ì—Ì«·" & CHR(13)
                            Msg = Msg + "·« Ì„ﬂ‰  ”ÃÌ· √ﬂÀ— „‰ ﬁÿ⁄… Ê«Õœ…" & CHR(13)
                            Msg = Msg + "Ì„ﬂ‰ﬂ  ”ÃÌ· ﬂ· ﬁÿ⁄… ›Ì ’› „” ﬁ·"
                        Else
                        Msg = "Item " & CHR(13)
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            Msg = Msg + "have Serial" & CHR(13)
                            Msg = Msg + "more item " & CHR(13)
                            Msg = Msg + "Ì„ﬂ‰ﬂ  ”ÃÌ· ﬂ· ﬁÿ⁄… ›Ì ’› „” ﬁ·"
                        
                        End If
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Count")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Count")
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                            'ÿÊ· «·—ﬁ„ «·”Ì—Ì«· ·«Ì“Ìœ ⁄‰ 50 Õ—›
                        ElseIf Len(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) > 50 Then
                            Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                            Msg = Msg & StrCurrentItemName & CHR(13)
                            Msg = Msg + "€Ì— „ﬁ»Ê·" & CHR(13)
                            Msg = Msg + "ÌÃ» √·« Ì“Ìœ ÿÊ· «·”Ì—Ì«· ⁄‰ 50 Õ—› √Ê —ﬁ„"
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Grid.Row = LngRowNum
                            Grid.Col = Grid.ColIndex("Serial")
                            Grid.ShowCell LngRowNum, Grid.ColIndex("Serial")
                            Grid.SetFocus
                            Screen.MousePointer = vbDefault
                            Exit Function
                        ElseIf .ColIndex("guaranteeTime") <> -1 Then

                            If .TextMatrix(LngRowNum, .ColIndex("guaranteeTime")) = "" Then
                            .TextMatrix(LngRowNum, .ColIndex("guaranteeTime")) = 1
                                If Not (IsNumeric(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) < 1 Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) > 36 Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) <> Int(val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Then
                                    Msg = "„œ… «·÷„«‰ «·Œ«’… »«·’‰›" & CHR(13)
                                    Msg = Msg + StrCurrentItemName & CHR(13)
                                    Msg = Msg + "€Ì— ’«·Õ…" & CHR(13)
                                    Msg = Msg + "√œŒ· „œ…   —«ÊÕ »Ì‰ 1 : 36 ‘Â—" & CHR(13)
                                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                    Msg = " „ «œŒ«· «·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                    Msg = Msg + StrCurrentItemName & CHR(13)
                    Msg = Msg + "›Ï «·”ÿ—:" & LngRowNum
                    Msg = Msg + "ﬁ»· –·ﬂ «·Ì «·’‰›" & CHR(13)
                    Msg = Msg + StrTemp & CHR(13)
                    Msg = Msg + "›Ï «·”ÿ—:" & LngSearchResault
                    Msg = Msg + " √ﬂœ „‰ «·”Ì—Ì«· «·Œ«’ »«·√’‰«›"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
                        LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
                        Set RsRepeat = New ADODB.Recordset

                        If Me.TxtModFlag.Text = "N" Then
                            Set RsRepeat = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), , , , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                        ElseIf Me.TxtModFlag.Text = "E" Then
                
                If Me.GridTrans = PurchaseTransaction Then
                  Set RsRepeat = New ADODB.Recordset
                            Set RsRepeat = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), , val(Me.TxtInvID.Text), , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                  Else
                  
                  Set RsRepeat = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), , val(Me.TxtInvID.Text), , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                  
                  
                  End If
                            
                  '      Print TxtNots.Text
                        
                        
                        End If

                        If Not (RsRepeat.EOF Or RsRepeat.BOF) Then
                            If RsRepeat("totalQty").value > 0 Then
                                Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                                Msg = Msg + StrCurrentItemName & CHR(13)
                                Msg = Msg + "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & CHR(13)
                                Msg = Msg + "„ÊÃÊœ ‰›” «·”Ì—Ì«· ·œÌ‰« ›Ï «·„Œ“‰ "

                                If Not StoreName Is Nothing Then
                                    Msg = Msg & "[" & StoreName.Text & "]"
                                End If

                                Msg = Msg & CHR(13) & "------------------------------------"
                                Msg = Msg & CHR(13) & "Õ«Ê· «‰  ” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ ”Ì—Ì«· ·„⁄—›… «·„“Ìœ"
                                Msg = Msg & CHR(13) & "Â·  —Ìœ «‰ ÌﬁÊ„ «·»—‰«„Ã »«·»ÕÀ ·ﬂ ⁄‰  Â–« «·”Ì—Ì«· «·¬‰...!!!"
                                IntRes = MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

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
                
                        If val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType"))) = 0 Then '·«»œ Ê«‰ ÌﬂÊ‰ «·’‰› ”·⁄…
                            If Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")) <> "" Then
                                '-------------------------------------------------------------------
                                '«·»Ì«‰«  «·√”«”Ì… —ﬁ„ «·’‰› Ê «”„ «·’‰› Ê «·”Ì—Ì«·
                                LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
                                StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
                                StrCurrentItemCode = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Code"))
                                LngColorID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ColorID")))
                                StrItemSize = Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemSize"))
                                LngClassId = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ClassId")))

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
                                    Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), Me.DtpBillDate.value, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                ElseIf Me.TxtModFlag.Text = "E" Then
                                    '·«ÕŸ «‰Â ›Ï Õ«·… «· ⁄œÌ· ·«»œ „‰ ≈—”«· —ﬁ„ «·›« Ê—… ··œ«·…
                        
                                    If (Me.GridTrans = InvoiceTransaction) Or (Me.GridTrans = Returntransaction) Then
                      
                                        Set RsTest = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), Me.DtpBillDate.value, val(Me.TxtNots.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                        
                                        'Set RsTest = GetItemQuantityStock(LngItemID, val(Me.storename.BoundText), _
                                        'Me.DtpBillDate.value,  Val(Me.TxtInvID.text), , , StrItemSerial, True, LngColorID, StrItemSize)
                                    Else
                        
                                        Set RsTest = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), Me.DtpBillDate.value, val(Me.TxtInvID.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                        
                                    End If
                        
                                End If

                                'Stop here
                                Screen.MousePointer = vbDefault
                    
                                If RsTest.EOF Or RsTest.BOF Then
                                    Msg = "«·’‰› : " & StrCurrentItemName & CHR(13)
                                    Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & CHR(13)

                                    If StrItemSerial <> "" Then
                                        Msg = Msg & "«·”Ì—Ì«· : " & StrItemSerial & CHR(13)
                                    End If
                        
                                    Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & CHR(13)
                                    Msg = Msg & "·« ÊÃœ „‰Â «Ï ﬂ„Ì… ›Ï «·„Œ“‰ : " & Me.StoreName.Text & CHR(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & CHR(13)
                                    Msg = Msg & "--------------------------" & CHR(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & CHR(13) & ""
                                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                                    FocusGridOnCell LngRowNum, .ColIndex("Name")
                                    TerminateRecordset RsTest
                                    Exit Function
                                ElseIf (Round(RsTest("totalqty").value, 2) < Round(DblOutQty, 2)) And StrItemSerial = "" Then
                                    Msg = "«·’‰› : " & StrCurrentItemName & CHR(13)
                                    Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & CHR(13)
                                    Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & CHR(13)
                                    Msg = Msg & "≈Ã„«·Ï «·ﬂ„Ì… «·„ÊÃÊœ… „‰Â ›Ï «·„Œ“‰ : " & Me.StoreName.Text & CHR(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & CHR(13)
                                    Msg = Msg & "ÂÏ : " & RsTest("totalqty").value & CHR(13)
                                    Msg = Msg & "«·ﬂ„Ì… «·„—«œ ’—›Â« : " & DblOutQty & CHR(13)
                                    Msg = Msg & "·«Ì„ﬂ‰ «·”„«Õ »Œ—ÊÃ Â–Â «·ﬂ„Ì….  " & CHR(13)
                                    Msg = Msg & "--------------------------" & CHR(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & CHR(13) & ""
                                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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




Private Function CheckTotalQty(LngRow As Long, _
                               Optional LngColorID1 As Long, _
                               Optional StrItemSize1 As String, _
                               Optional LngClassId1 As Long, _
                               Optional StrItemSerial1 As String) As Double
    Dim i As Integer
    Dim LngFoundGridRow As Long
    Dim Dbl_OutQtyBySmallUnit As Double

    Dim SmalltUnit As Double
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
    Dim Quantity As Double
    Dim StrSQL As String
    Dim RsUnitData  As New ADODB.Recordset
    Dim LngColorID As Long
    Dim LngClassId As Long
    Dim StrItemSize As String
    Dim StrItemSerial As String

    If LngRow = 0 Then Exit Function

    With Me.Grid
        LngFoundGridRow = 0

        For i = 1 To .Rows - 1
            LngFoundGridRow = .FindRow(.TextMatrix(LngRow, .ColIndex("Code")), i, .ColIndex("Code"), True, True)

            If LngFoundGridRow > 0 Then
                i = LngFoundGridRow
            
                LngCurItemID = .TextMatrix(LngFoundGridRow, .ColIndex("Code"))
                LngUnitID = val(.Cell(flexcpData, LngFoundGridRow, .ColIndex("UnitID")))
                DblQty = val(.TextMatrix(LngFoundGridRow, .ColIndex("Count")))

                LngColorID = val(Grid.TextMatrix(LngFoundGridRow, Grid.ColIndex("ColorID")))
                StrItemSize = Grid.TextMatrix(LngFoundGridRow, Grid.ColIndex("ItemSize"))
                LngClassId = val(Grid.TextMatrix(LngFoundGridRow, Grid.ColIndex("ClassId")))

                StrItemSerial = Grid.TextMatrix(LngFoundGridRow, Grid.ColIndex("Serial"))
      
                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    SmalltUnit = RsUnitData("UnitFactor").value

                    If StrItemSerial = StrItemSerial1 And LngColorID = LngColorID1 And StrItemSize = StrItemSize1 And LngClassId = LngClassId1 Then
                        Quantity = SmalltUnit * val(.TextMatrix(LngFoundGridRow, .ColIndex("Count")))
                    Else
                        Quantity = 0
                    End If

                    ' Dbl_OutQtyBySmallUnit = Quantity
                    ' Exit For
                End If
            
                Dbl_OutQtyBySmallUnit = Dbl_OutQtyBySmallUnit + Quantity 'Val(.TextMatrix(LngFoundGridRow, .ColIndex("Count")))
            ElseIf LngFoundGridRow = -1 Then        'did not found the item entered before
                Exit For
            End If

        Next i

    End With

    CheckTotalQty = Dbl_OutQtyBySmallUnit
End Function

Private Sub FocusGridOnCell(LngRow As Long, _
                            LngCol As Long)
    On Local Error GoTo ErrTrap

    With Me.Grid
        .Row = LngRow
        .Col = LngCol
        .ShowCell LngRow, LngCol
        .SetFocus
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub ProductionSearchFun(Row As Long, _
                          Col As Long, _
                          Item_ID As Long)
    Dim Frm As Order_no_search2
    Set Frm = New Order_no_search2
 
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
        Frm.lbl(12).Caption = (Row + 1)
 
     


     
        Set Frm.NewGrid = Me.Grid
        Frm.LngRow = Row
       Frm.RetrunType = 8
       Frm.show
       Frm.lblitemid.Caption = Item_ID
        'Frm.LngCol = Col
    
        '' Frm.XPDtbBill.value = Date
 
    End With

    Frm.show vbModal
End Sub

Private Sub ItemSearchFun(Row As Long, _
                          Col As Long, _
                          Item_ID As Long)
    Dim Frm As FrmItemSearch1
    Set Frm = New FrmItemSearch1
 
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
        Frm.lbl(12).Caption = (Row + 1)
        '  ' Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        ' '  Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        'Set Frm.NewGrid.Grid = FG1
        ' Set Frm.NewGrid.Grid = New Grid
        '    Set Frm.NewGrid = New Frm.NewGrid
        ' NewGrid.GridTrans = Me.GridTrans
     
        Set Frm.NewGrid = Me.Grid
        Frm.LngRow = Row
    
        'Frm.LngCol = Col
    
        '' Frm.XPDtbBill.value = Date
 
    End With

    Frm.show vbModal
End Sub
  
Private Sub ItemLotNoFun(Row As Long, _
                         Col As Long, _
                         Item_ID As Long, _
                         StoreId As Long, _
                         rdate As Date, Optional Indxx As Integer)
       If SystemOptions.AllowExperDateFIFO = False Then
    Dim Frm As ItemLOtNO
    Set Frm = New ItemLOtNO
 
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
    If Row <> 0 Then
    Frm.lbl(0).Visible = True
       Frm.lbl(3).Visible = True
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
     Else
        Frm.lbl(3).Visible = False
            Frm.lbl(0).Visible = False
        'Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.DCboItemCode
        Frm.lbl(5).Caption = Me.DCboItemName.Text
     End If
     
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        Frm.XPDtbBill.value = Date
    
        '***********************************************
 
        Dim i As Integer
        Dim rs As ADODB.Recordset
        Dim My_SQL As String

        Set rs = New ADODB.Recordset
        My_SQL = " SELECT     TOP 100 PERCENT  dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) "
        My_SQL = My_SQL & " AS Qty ,dbo.Transaction_Details.ProductionDate "
        My_SQL = My_SQL & " FROM         dbo.Transaction_Details INNER JOIN"
        My_SQL = My_SQL & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
        My_SQL = My_SQL & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
        My_SQL = My_SQL & " Where (dbo.TransactionTypes.StockEffect <> 0) and (dbo.Transactions.StoreID = " & StoreId & ") And (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
        My_SQL = My_SQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(rdate, True) & ""

        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Then
            My_SQL = My_SQL + " and dbo.Transactions.Transaction_ID<>" & val(Me.TxtNots)
        Else
            My_SQL = My_SQL + " and dbo.Transactions.Transaction_ID<>" & val(Me.TxtInvID)
        End If
 
        My_SQL = My_SQL & " GROUP BY dbo.Transaction_Details.LotNO, dbo.Transaction_Details.ExpiryDate,dbo.Transaction_Details.ProductionDate"
        My_SQL = My_SQL & " ORDER BY dbo.Transaction_Details.ExpiryDate"

        rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

Frm.Indxx = Indxx
        With Frm
            With Frm.QtyGrid
                .Rows = 2
                .Clear flexClearScrollable

                If rs.RecordCount > 0 Then
                    .Rows = rs.RecordCount + 1
                    rs.MoveFirst

                    For i = 1 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Ser")) = i
                     
                        .TextMatrix(i, .ColIndex("LotNO")) = IIf(IsNull(rs.Fields("LotNO").value), "", rs.Fields("LotNO").value)
                       
                        .TextMatrix(i, .ColIndex("ExpiryDate")) = IIf(IsNull(rs.Fields("ExpiryDate").value), "", rs.Fields("ExpiryDate").value)
                    
                        .TextMatrix(i, .ColIndex("ProductionDate")) = IIf(IsNull(rs.Fields("ProductionDate").value), "", rs.Fields("ProductionDate").value)
                       
                        .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs.Fields("Qty").value), "", rs.Fields("Qty").value)
                    
                        rs.MoveNext
                    Next

                    rs.Close
                End If

                .RowHeight(-1) = 300
            End With
        End With
 
        '**************************************************
        If Me.TxtModFlag.Text = "R" Then
            Frm.TxtComment.locked = True
            Frm.CmdOk.Enabled = False
        End If
    End With
    Frm.show
    End If
End Sub
  
   
 Private Sub ItemGranteeData(Row As Long, _
                            Col As Long, _
                            Optional calcExpiryDate As Boolean)
    Dim Frm As FRMGranteeData
    Set Frm = New FRMGranteeData
    Dim Item_ID As Integer
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Item_ID = Me.Grid.TextMatrix(Row, .ColIndex("Code"))

 

 
    
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        Frm.GranteeStartDate.value = Date
    
        If 1 = 1 Then
            If Me.TxtModFlag.Text = "R" Then
                Frm.TxtComment.locked = True
                Frm.CmdOk.Enabled = False
            End If

            If Me.Grid.ColIndex("GranteeEndDate") <> -1 Then
                Frm.GranteeEndDate.value = IIf(Grid.TextMatrix(Row, Grid.ColIndex("GranteeEndDate")) = "", Date, Grid.TextMatrix(Row, Grid.ColIndex("GranteeEndDate")))
            End If
        
            If Me.Grid.ColIndex("GranteeStartDate") <> -1 Then
                Frm.GranteeStartDate.value = IIf(Grid.TextMatrix(Row, Grid.ColIndex("GranteeStartDate")) = "", Date, Grid.TextMatrix(Row, Grid.ColIndex("GranteeStartDate")))
            End If
        
            If Me.Grid.ColIndex("RegularMaintenancedates") <> -1 Then
                Frm.AllDate = IIf(Grid.TextMatrix(Row, Grid.ColIndex("RegularMaintenancedates")) = "", Date, Grid.TextMatrix(Row, Grid.ColIndex("RegularMaintenancedates")))
            End If
        
        
            If Me.Grid.ColIndex("RegularMaintenanceIDS") <> -1 Then
                Frm.AllIDS = IIf(Grid.TextMatrix(Row, Grid.ColIndex("RegularMaintenanceIDS")) = "", "", Grid.TextMatrix(Row, Grid.ColIndex("RegularMaintenanceIDS")))
            End If
            
            
            
            If Me.Grid.ColIndex("guaranteeTime") <> -1 Then
                Frm.txtvlaue = IIf(Grid.TextMatrix(Row, Grid.ColIndex("guaranteeTime")) = "", 12, Grid.TextMatrix(Row, Grid.ColIndex("guaranteeTime")))
            End If
        
            If Me.Grid.ColIndex("GranteeType") <> -1 Then
                 
                If val(Grid.TextMatrix(Row, Grid.ColIndex("GranteeType"))) = 1 Then
                    Frm.GranteeTypeopt(1).value = True
                Else
                    Frm.GranteeTypeopt(0).value = True
                End If
        
            End If

            Frm.txtvlaue_Change
            Frm.FillGridWithData
        
        End If

    End With

    Frm.show vbModal
End Sub
Private Sub ItemGoldData(Row As Long, _
                            Col As Long, _
                            Optional calcExpiryDate As Boolean)
   ' Dim Frm As FrmGoldDetaiks
   ' Set Frm = New FrmGoldDetaiks
   ' Dim Item_ID As Integer
   ' Dim GroupID As Integer
   ' Dim ExpiryValue As Integer
   ' Dim ExpiryType As Integer
'
    'With Me.Grid
'        Frm.lbl(3).Caption = (Row + 1)
'        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
'        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
'        Item_ID = Me.Grid.TextMatrix(Row, .ColIndex("Code"))
'
'

 
    
        'Set Frm.Fg = Me.Grid
        'Frm.LngRow = Row
        'Frm.LngCol = Col
    
        'Frm.GranteeStartDate.value = Date
    
        'If 1 = 1 Then
        '    If Me.TxtModFlag.Text = "R" Then
        '        Frm.TxtComment.locked = True
        '        Frm.CmdOk.Enabled = False
        '        Frm.fgCameo.Enabled = False
        '         Frm.txtWages.Enabled = False
        '
        '    End If

        '    If Me.Grid.ColIndex("GranteeEndDate") <> -1 Then
        '        Frm.GranteeEndDate.value = IIf(Grid.TextMatrix(Row, Grid.ColIndex("GranteeEndDate")) = "", Date, Grid.TextMatrix(Row, Grid.ColIndex("GranteeEndDate")))
        '    End If
        
        '    If Me.Grid.ColIndex("Wages") <> -1 Then
        '        Frm.txtWages = IIf(Grid.TextMatrix(Row, Grid.ColIndex("Wages")) = "", "", Grid.TextMatrix(Row, Grid.ColIndex("Wages")))
        '    End If
        
        '
        '    If Me.Grid.ColIndex("GoldDetails") <> -1 Then
        '        Frm.AllIDS = IIf(Grid.TextMatrix(Row, Grid.ColIndex("GoldDetails")) = "", "", Grid.TextMatrix(Row, Grid.ColIndex("GoldDetails")))
        '    End If
            
            
            
        '    If Me.Grid.ColIndex("guaranteeTime") <> -1 Then
        '        Frm.txtvlaue = IIf(Grid.TextMatrix(Row, Grid.ColIndex("guaranteeTime")) = "", 12, Grid.TextMatrix(Row, Grid.ColIndex("guaranteeTime")))
        '    End If
        
        '    If Me.Grid.ColIndex("GranteeType") <> -1 Then
                 
        '        If val(Grid.TextMatrix(Row, Grid.ColIndex("GranteeType"))) = 1 Then
        '            Frm.GranteeTypeopt(1).value = True
        '        Else
        '            Frm.GranteeTypeopt(0).value = True
        '        End If
        
        '    End If

        '    Frm.txtvlaue_Change
        '    Frm.FillGridWithData
        
        'End If

    'End With

    'Frm.show vbModal
End Sub

 
 Private Sub ItemDetailsNewIgia(Row As Long, _
                            Col As Long, _
                            Optional calcExpiryDate As Boolean)
    Dim Frm As FrmItemsDetails
    Set Frm = FrmItemsDetails
    
    
    Dim Item_ID As Variant
    Dim GroupID As Integer
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer
'*********************************khaled******************************
    Frm.TxtModFlagDet = Me.TxtModFlag.Text
If Not Me.TxtInvID Is Nothing Then
    Frm.TxtInvIDDet = Me.TxtInvID.Text
 End If
    Frm.GridTransDet = Me.GridTrans
'*********************************khaled******************************


    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Item_ID = val(Me.Grid.TextMatrix(Row, .ColIndex("Code")))

 Frm.TxtUnitID = .Cell(flexcpData, .Row, .ColIndex("UnitID"))
Frm.TxtItemID.Text = Item_ID
     FrmItemsDetails.txtSizeId = .TextMatrix(Row, .ColIndex("itemsize"))
        '    FrmItemsDetails.txtSizeName = .TextMatrix(Row, .ColIndex("itemsize"))
            
            
             FrmItemsDetails.txtClassId = .TextMatrix(.Row, .ColIndex("ClassId"))
      '      FrmItemsDetails.txtClassName = .TextMatrix(.Row, .ColIndex("ClassId"))
            
            
             FrmItemsDetails.txtColorID = .TextMatrix(.Row, .ColIndex("ColorID"))
            'FrmItemsDetails.txtColorName = .TextMatrix(.Row, .ColIndex("ColorID"))
            Dim UnitName As String
            Dim sizename As String
             Dim colorname As String
             Dim classname As String
            
            
              GetiItemsNewDetails val(Frm.TxtUnitID), val(FrmItemsDetails.txtSizeId) _
  , val(FrmItemsDetails.txtColorID), val(FrmItemsDetails.txtClassId), UnitName, sizename, colorname, classname
   FrmItemsDetails.txtSizeName = sizename
   FrmItemsDetails.txtColorName = colorname
   FrmItemsDetails.txtClassName = classname
   
    
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        'Frm.GranteeStartDate.value = Date
    
        If 1 = 1 Then
            If Me.TxtModFlag.Text = "R" Then
             '  Frm.TxtComment.locked = True
              Frm.Cmd(2).Enabled = False
             Frm.Grid.Enabled = False
             Frm.TxtxreateSerial.Enabled = False
             '    Frm.txtWages.Enabled = False
                
            End If

 
        
            
            If Me.Grid.ColIndex("ItemsDetailsNewidea") <> -1 Then
                Frm.AllIDS = IIf(Grid.TextMatrix(Row, Grid.ColIndex("ItemsDetailsNewidea")) = "", "", Grid.TextMatrix(Row, Grid.ColIndex("ItemsDetailsNewidea")))
            End If
            
            
            
 
            Frm.FillGridWithData
        
        End If

    End With

    Frm.show
   
End Sub


Private Sub ItemProductionDate(Row As Long, _
                               Col As Long, _
                               Optional calcExpiryDate As Boolean, Optional datetype As Integer)
    Dim Frm As ItemProductionDate
    Set Frm = New ItemProductionDate
    Dim Item_ID As Long
    Dim GroupID As Double
    Dim ExpiryValue As Integer
    Dim ExpiryType As Integer

    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Item_ID = Me.Grid.TextMatrix(Row, .ColIndex("Code"))
        GetItemData Item_ID, , , GroupID
        GetGroupData GroupID, , , , , "Groups", ExpiryType, ExpiryValue


 
 GetItemIDExpiry Item_ID, ExpiryType, ExpiryValue

 

        If Me.Grid.ColIndex("ExpiryDate") <> -1 And calcExpiryDate = True Then
            Frm.lbl(7).Caption = ExpiryValue
                
            Frm.lbl(8).Caption = ExpiryType
        ElseIf calcExpiryDate = False Or Me.Grid.ColIndex("ExpiryDate") = -1 Then
            Frm.lbl(7).Caption = -1
            
            Frm.lbl(8).Caption = -1
        End If
    
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        Frm.XPDtbBill.value = Date
    
        If Me.TxtModFlag.Text = "R" Then
            Frm.TxtComment.locked = True
            Frm.CmdOk.Enabled = False
        End If

    End With


    
    If datetype = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                 Frm.lbl(6).Caption = " «—ÌŒ «·«‰ «Ã"
                 Else
                 Frm.lbl(6).Caption = "Production Date"
                 End If
    
    Else
    
                    If SystemOptions.UserInterface = ArabicInterface Then
                 Frm.lbl(6).Caption = " «—ÌŒ «·«‰ Â«¡"
                 Else
                 Frm.lbl(6).Caption = "Expiry Date"
                 End If

    End If
        Frm.show vbModal
End Sub

Private Sub ItemSSErials(Row As Long, _
                         Col As Long)
    Dim Frm As FrmGridAddItemSerials
    Set Frm = New FrmGridAddItemSerials

    If Not Me.StoreName Is Nothing Then
        FrmGridAddItemSerials.DcboStores.BoundText = val(Me.StoreName.BoundText)
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«Œ — „Œ“‰ «Ê·«", vbCritical
        Else
            MsgBox "Select Store Firstly", vbCritical
        End If
        
        Exit Sub
    End If
                
    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Frm.ItemID.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Name"))
        
        Frm.Vatyo.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Vatyo"))
        Frm.Vat.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Name"))
        
                 Frm.lblunitid.Caption = .Cell(flexcpData, .Row, .ColIndex("UnitID"))
            Frm.lblunitname.Caption = .TextMatrix(.Row, .ColIndex("UnitId"))
            
Frm.TxtPrice.Text = .TextMatrix(.Row, .ColIndex("Price"))


        Set Frm.FG = Me.Grid
    
        Frm.LngRow = Row
        Frm.LngCol = Col
        '   Frm.TxtComment.text = Trim$(Me.Grid.TextMatrix(Row, Col))
        Frm.TxtComment.Text = ""
   
        Dim i As Integer

        For i = 1 To Grid.Rows - 1

            If Me.Grid.Cell(flexcpTextDisplay, i, .ColIndex("Code")) = Frm.lbl(4).Caption Then
                If i < Grid.Rows - 1 Then
                    Frm.TxtComment.Text = Frm.TxtComment.Text & Trim$(Me.Grid.TextMatrix(i, Col)) & ","
                Else
                    Frm.TxtComment.Text = Frm.TxtComment.Text & Trim$(Me.Grid.TextMatrix(i, Col))
                End If
                            
            End If

        Next i
  If Len(Frm.TxtComment.Text) > 1 Then
  Frm.TxtComment.Text = mId(Frm.TxtComment.Text, 1, Len(Frm.TxtComment.Text) - 1)
  End If
        If Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = 7 Then
            Frm.FG1.Enabled = True
            Frm.TxtComment.Enabled = False

            If Me.GridTrans = InventoryOut Then
                Frm.FillSerials val(Frm.ItemID.Caption), val(val(Me.StoreName.BoundText)), val(Me.TxtInvID.Text)
            Else
                Frm.FillSerials val(Frm.ItemID.Caption), val(val(Me.StoreName.BoundText)), val(Me.TxtInvID.Text)

            End If
        
        Else
            Frm.FG1.Enabled = False
            FrmGridAddItemSerials.TxtComment.Enabled = True

        End If
    
        If Me.TxtModFlag.Text = "R" Then
            Frm.TxtComment.locked = True
            Frm.CmdOk.Enabled = False
        End If

    End With

     
     Frm.show vbModal
End Sub

Private Sub ItemComment(Row As Long, _
                        Col As Long)
    Dim Frm As FrmGridAddItemComment
    Set Frm = New FrmGridAddItemComment

    With Me.Grid
        Frm.lbl(3).Caption = (Row + 1)
        Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
        Frm.TxtComment.Text = Trim$(Me.Grid.TextMatrix(Row, Col))

        If Me.TxtModFlag.Text = "R" Then
            Frm.TxtComment.locked = True
            Frm.CmdOk.Enabled = False
        End If

    End With

    Frm.show
End Sub

Private Sub PopUpSerialList(LngItemID As Long, _
                            Optional LngRow As Long = 0, _
                            Optional LngCol As Long = 0)

  '  Dim Frm As FrmSerialList
  '  Dim LngStoreID As Long
  '  LngStoreID = val(val(Me.storename.BoundText))
'
'    If LngItemID = 0 Or LngStoreID = 0 Then
'        Exit Sub
'    End If
'
'    Set Frm = New FrmSerialList
'    Frm.RetrunType = 0
'    Frm.RetrunGridOnOneline = False
'    Set Frm.m_TextBox = Me.TxtSerial
'    Frm.GetData LngItemID, LngStoreID
'    '------------------------------------------
'    Frm.FgCol = LngCol
'    Frm.FgRow = LngRow
'    '------------------------------------------
'    Set Frm.xGrid = Me
'    Frm.show vbModal

End Sub
 Private Sub ClearTopDataEntry()
   If ReturnTyp <> 10 Then
    DCboItemCode.BoundText = ""
    DCboItemName.BoundText = ""
    CboItemCase.ListIndex = 0
    End If
              If SystemOptions.AllowItemsShortName = True Then
                                    If Not TxtShortName Is Nothing Then
                                    TxtShortName.Text = ""
                                        TxtShortName.SetFocus
                            End If
           End If
           
    If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
        TxtSerial.Text = ""
    End If

    TxtQuantity.Text = ""
    TxtPrice.Text = ""
    CboItemCase.ListIndex = 0
End Sub

Private Sub TxtSerial_Change()
DcboItemName_Change
End Sub

Private Sub TxtServiceTax_Change()

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
        Me.Calculate 1, , , True
    End If

End Sub

Private Sub TxtServiceTax_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtServiceTax.Text, 0)
End Sub

Private Sub TxtStampTax_Change()

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
        Me.Calculate 1, , , True
    End If

End Sub

Private Sub TxtStampTax_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtStampTax.Text, 0)
End Sub

Private Sub TxtTaxValue_Change()

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
        Me.Calculate 1, , , True
    End If

End Sub

Private Sub TxtTaxValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTaxValue.Text, 0)
End Sub

Public Function GetItemsTotal(m_ItemType As ItemsType, Optional ByRef ItemCount As Double) As Double
    Dim DblTemp As Double
    Dim RowNum As Long
    Dim Msg  As String
     On Local Error GoTo ErrTrap
Dim linetotl As Double
ItemCount = 0
    With Me.Grid

        If m_ItemType = ItemsAllType Then

            For RowNum = .FixedRows To .Rows - 1
            ItemCount = ItemCount + 1
            linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
            
            Next RowNum

        ElseIf m_ItemType = ItemsGoodType Then

            For RowNum = .FixedRows To .Rows - 1

                If val(.TextMatrix(RowNum, .ColIndex("ItemType"))) = 0 And val(.TextMatrix(RowNum, .ColIndex("Name"))) <> 0 Then
                ItemCount = ItemCount + 1
                     linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
                    
                End If

            Next RowNum

        ElseIf m_ItemType = ItemsServiceType Then

            For RowNum = .FixedRows To .Rows - 1

                If val(.TextMatrix(RowNum, .ColIndex("ItemType"))) = 1 And val(.TextMatrix(RowNum, .ColIndex("Name"))) <> 0 Then
                ItemCount = ItemCount + 1
                          linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
                        
                End If

            Next RowNum

        ElseIf m_ItemType = 5 Then

            For RowNum = .FixedRows To .Rows - 1
            ItemCount = ItemCount + 1
                   linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
                    
            Next RowNum

        End If
    
    End With

    GetItemsTotal = DblTemp
    Exit Function
ErrTrap:
    Msg = "ÕœÀ Œÿ« "
    GetItemsTotal = DblTemp
End Function

Public Function GetItemsCostTotal() As Double
    Dim DblTemp As Double
    Dim RowNum As Long
    Dim Msg  As String
    Dim DblRowCostTotal As Double

    On Local Error GoTo ErrTrap

    With Me.Grid

        For RowNum = .FixedRows To .Rows - 1
            DblRowCostTotal = val(.TextMatrix(RowNum, .ColIndex("ItemCostPrice"))) * val(.TextMatrix(RowNum, .ColIndex("Count")))
            DblTemp = DblTemp + DblRowCostTotal
        Next RowNum

    End With

    GetItemsCostTotal = DblTemp
    Exit Function
ErrTrap:
    Msg = "ÕœÀ Œÿ« "
    GetItemsCostTotal = 0
End Function

Public Property Get RestrictedAssbliedItemID() As Long
    RestrictedAssbliedItemID = m_RestrictedAssbliedItemID
End Property

Public Property Let RestrictedAssbliedItemID(ByVal vNewValue As Long)
    m_RestrictedAssbliedItemID = vNewValue
End Property

Public Property Get WorkOrderID() As Long
    WorkOrderID = m_WorkOrderID
End Property

Public Property Let WorkOrderID(ByVal vNewValue As Long)
    m_WorkOrderID = vNewValue
End Property

Private Function CheckRestrictedItemQty(LngWorkOrderID As Long, _
                                        LngItemID As Long) As Boolean

    Dim Msg As String
    Dim i As Long
    Dim RsOutQty As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim SngWillGoOutQty As Single
    Dim SngMaxOutQty As Single
    Dim SngOutQty As Single
    Dim LngSearchItemID As Long
    Dim RsUserOptions As ADODB.Recordset
    Dim IntAllowMaxOverOut As Integer
    Dim IntRes As Integer
    Dim BolThereIsOutBefore As Boolean

    'IntAllowMaxOverOut=-1 NOT  Set
    'IntAllowMaxOverOut=0 NOT Allow
    'IntAllowMaxOverOut=1 Allowed

    StrSQL = "SELECT SUM(Qty) AS SumQty, ItemID, ItemCode, ItemName, StoreID, StoreName," & "GroupID, ColorID, ColorName, ItemSize "
    StrSQL = StrSQL + " From "
    StrSQL = StrSQL + "( "
    StrSQL = StrSQL + " SELECT TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName, Transactions.StoreID, TblStore.StoreName, TblItems.GroupID,"
    StrSQL = StrSQL + " Transaction_Details.ColorID, TblItemsColors.ColorName, Transaction_Details.ItemSize,"
    StrSQL = StrSQL + " Qty = CASE WHEN Transactions.Transaction_Type = 18 THEN Transaction_Details.Quantity * 1"
    StrSQL = StrSQL + " ELSE 0 END"
    StrSQL = StrSQL + " FROM         dbo.TblStore INNER JOIN"
    StrSQL = StrSQL + " dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
    StrSQL = StrSQL + " dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " Where dbo.Transactions.WorkOrderID=" & LngWorkOrderID & ")"
    StrSQL = StrSQL + " XTable"
    StrSQL = StrSQL + " GROUP BY ItemID, ItemCode, ItemName,StoreID,StoreName,GroupID,ColorID,ColorName,ItemSize"
    Set RsOutQty = New ADODB.Recordset
    RsOutQty.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsOutQty.BOF Or RsOutQty.EOF Then
        '·„ Ì „ ’—› «Ï ﬂ„Ì«  ⁄·Ï «„— «·‘€· Â–«
        BolThereIsOutBefore = False
    Else
        BolThereIsOutBefore = True
        'CheckRestrictedItemQty = True
        'Exit Function
    End If

    '-----------------------------------------------------------------------------
    StrSQL = "Select  * From TblUsers Where UserID=" & user_id
    Set RsUserOptions = New ADODB.Recordset
    RsUserOptions.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUserOptions.BOF Or RsUserOptions.EOF) Then
        If IsNull(RsUserOptions("AllowOverMax").value) = True Then
            IntAllowMaxOverOut = -1
        ElseIf RsUserOptions("AllowOverMax").value = 0 Then
            IntAllowMaxOverOut = 0
        ElseIf RsUserOptions("AllowOverMax").value = 1 Then
            IntAllowMaxOverOut = 1
        End If
    End If

    '-----------------------------------------------------------------------------
    StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty,dbo.TblItemsParts.TableId "
    StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID "
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid

        For i = 1 To .Rows - 1
            SngOutQty = 0
            SngMaxOutQty = 0
            SngWillGoOutQty = 0
        
            LngSearchItemID = val(.TextMatrix(i, .ColIndex("Name")))
            SngWillGoOutQty = val(.TextMatrix(i, .ColIndex("Count")))

            If BolThereIsOutBefore = True Then
                RsOutQty.MoveFirst
                RsOutQty.find "ItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

                If Not (RsOutQty.BOF Or RsOutQty.EOF) Then
                    SngOutQty = IIf(IsNull(RsOutQty("SumQty").value), 0, RsOutQty("SumQty").value)
                Else
                    SngOutQty = 0
                End If

            Else
                SngOutQty = 0
            End If

            rs.MoveFirst
            rs.find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

            If Not (rs.BOF Or rs.EOF) Then
                SngMaxOutQty = rs("PartItemQty").value

                If (SngWillGoOutQty + SngOutQty) > SngMaxOutQty Then
                    Msg = "«·’‰›..." & rs("ItemName").value
                    Msg = Msg & CHR(13) & "√ﬁ’Ì ﬂ„Ì… „‰Â ›Ï «·’‰› «·„Ã„⁄ :" & SngMaxOutQty

                    If IntAllowMaxOverOut = -1 Then
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…: ·„ Ìﬁ„ „œÌ— «·‰Ÿ«„ »≈⁄œ«œ ’·«ÕÌ« ﬂ ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                        Msg = Msg & CHR(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        CheckRestrictedItemQty = False
                        Exit Function
                    ElseIf IntAllowMaxOverOut = 0 Then
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…: ·Ì” ·ﬂ ’·«ÕÌ… ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                        Msg = Msg & CHR(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        CheckRestrictedItemQty = False
                        Exit Function
                    ElseIf IntAllowMaxOverOut = 1 Then
                        Msg = Msg & CHR(13) & "Â·  —Ìœ ’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄"
                        IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                        If IntRes = vbNo Then
                            CheckRestrictedItemQty = False
                            Exit Function
                        End If
                    End If
                End If
            End If

        Next i

    End With

    CheckRestrictedItemQty = True
End Function

Private Function CheckRestrictedItem() As Boolean

    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim LngSearchItemID As Long
    Dim StrTemp As String
    Dim i As Long
    Dim IntRes As Integer

    StrSQL = "SELECT dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty,dbo.TblItemsParts.TableId "
    StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID "
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & Me.RestrictedAssbliedItemID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "«·’‰› «·„Õœœ ’‰› €Ì— „Ã„⁄ »—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CheckRestrictedItem = False
        Exit Function
    End If

    With Me.Grid

        For i = .FixedRows To .Rows - 1
            LngSearchItemID = val(.TextMatrix(i, .ColIndex("Name")))

            If LngSearchItemID <> 0 Then
                rs.MoveFirst
                rs.find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

                If rs.BOF Or rs.EOF Then
                    StrTemp = "«·’‰› ..." & .Cell(flexcpTextDisplay, i, .ColIndex("Name"))
                    Msg = StrTemp
                    Msg = Msg & CHR(13) & "€Ì— „ÊÃÊœ ›Ï «·’‰› «·„Ã„⁄ «·„— »ÿ »Â «„— «·‘€·"
                    Msg = Msg & CHR(13) & ""
                    Msg = Msg & CHR(13) & "Â·  —Ìœ «·„ «»⁄… ›Ï Õ›Ÿ «„— «·‘€·...øøø"
                    IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                    If IntRes = vbNo Then
                        CheckRestrictedItem = False
                        Exit Function
                    End If
                End If
            End If

        Next i

        CheckRestrictedItem = True
        Exit Function
    End With

End Function

Private Sub PutDefCount(LngRow As Long)
    Dim RsUnitData As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
    Dim DblOrgQty As Double
    Dim StrSQL As String

    'Dim Dbl
    With Me.Grid
        LngItemID = val(.TextMatrix(LngRow, .ColIndex("Name")))
        LngUnitID = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
    
        If LngItemID = 0 Then Exit Sub
   
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
        StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
        StrSQL = StrSQL + " AND TblItemsUnits.UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        DblOrgQty = val(.TextMatrix(LngRow, .ColIndex("Count")))
    
        '    If MDIFrmMain.ActiveForm.name = "FrmSaleBill" Or MDIFrmMain.ActiveForm.name = "FrmReturnSalling" Or MDIFrmMain.ActiveForm.name = "FrmShowPrice" Or MDIFrmMain.ActiveForm.name = "FrmTemplate" Then
        '    .TextMatrix(LngRow, .ColIndex("Price")) = RsUnitData("UnitSalesPrice")
        '   Else
        '    .TextMatrix(LngRow, .ColIndex("Price")) = RsUnitData("UnitPurPrice")
        '    End If
   
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            '        .TextMatrix(LngRow, .ColIndex("defCount")) = DblOrgQty * RsUnitData("FactorByDefaultUnit").Value
            '        .TextMatrix(LngRow, .ColIndex("defCount")) = Format(.TextMatrix(LngRow, .ColIndex("DefCount")), "0.000")
        End If
    
    End With

    RsUnitData.Close
    Set RsUnitData = Nothing
End Sub

Public Sub DtpBillDate_Change()
 If Me.GridTrans <> NewGard And Me.GridTrans <> PurchaseTransaction Then
 If bassprofit = False Then
  updateProfit
 Else
  bassprofit = False
 End If
 
 End If
 If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowItemByCustomer = True Then
 Me.FillGrid , 1
 End If
SentTypeVAT
End Sub
'Public Sub TxtDiscount_Val_Change()
'SentTypeVAT
'End Sub
Public Sub CboDiscount_Type_Change()
SentTypeVAT
End Sub
 
 
Sub SentTypeVAT()

 If Me.TxtModFlag.Text <> "R" Then
   
Dim i As Long
         With Grid
              For i = 1 To .Rows - 1
    
             If Me.GridTrans = InvoiceTransaction Then
               CalculteValueAdded i, 21, , 1
             ElseIf Me.GridTrans = PurchaseTransaction Then
               CalculteValueAdded i, 22, , 1
             ElseIf Me.GridTrans = ReturnSalling Then
               CalculteValueAdded i, 9, , 1
             ElseIf Me.GridTrans = Returntransaction Then
               CalculteValueAdded i, 5, , 1
             ElseIf Me.GridTrans = purchaseOrderApproved Then
               CalculteValueAdded i, 29, , 1
             ElseIf Me.GridTrans = salespricelist Then
               CalculteValueAdded i, 42, , 1
             ElseIf Me.GridTrans = ShowPrice Then
               CalculteValueAdded i, 30, , 1
             End If
      
                Next i
                
               End With
      End If
      FillGridValuAddedAllItems
End Sub

 Private Sub Customer_Change()
Customer_Click (0)
End Sub
 
 Private Sub Customer_Click(Area As Integer)

If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowItemByCustomer = True Then
FillGrid , 1, False
End If
End Sub


