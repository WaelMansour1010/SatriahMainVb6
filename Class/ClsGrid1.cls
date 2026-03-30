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
Public WithEvents Grid As vsFlexGrid
Attribute Grid.VB_VarHelpID = -1
Public WithEvents TxtModFlag As VB.TextBox
Attribute TxtModFlag.VB_VarHelpID = -1
Public WithEvents GrdTBar As MSComctlLib.Toolbar
Attribute GrdTBar.VB_VarHelpID = -1
Public TxtTotal As TextBox
Attribute TxtTotal.VB_VarHelpID = -1
Public CboDiscount_Type As ComboBox
Public TxtDiscount_Val As TextBox
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
Public StoreName As DataCombo
Public LblInvProfit As Label
Public LblItemsCount As Label
Public DtpBillDate As DTPicker
Public TxtInvID As TextBox
Private FrmComment As FrmFlexQtyTip
Public WithEvents m_FrmItemsNewPrices As Form
Attribute m_FrmItemsNewPrices.VB_VarHelpID = -1
Public WithEvents LblTotalAll As Label
Attribute LblTotalAll.VB_VarHelpID = -1
Public WithEvents LblDiscountsTotal As Label
Attribute LblDiscountsTotal.VB_VarHelpID = -1
Public LblTaxSalesValue As Label
Public LblTaxAddValue As Label
Public LblTaxStampValue As Label
Public LblTaxServiceValue As Label

'«·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
Public WithEvents DCboItemCode As DataCombo  'ﬂÊœ «·’‰›
Attribute DCboItemCode.VB_VarHelpID = -1
Public WithEvents DcboItemName As DataCombo  '«”„ «·’‰›
Attribute DcboItemName.VB_VarHelpID = -1
Public WithEvents CboItemCase As ComboBox    'Õ«·… «·’‰›
Attribute CboItemCase.VB_VarHelpID = -1
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
Public Function FindItemInGrid(LngRow As Long, LngItemID As Long) As Long

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

Private Sub CmdAddData_Click()
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

If DCboItemCode.text = "" Then
    Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DCboItemCode.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If
If DcboItemName.text = "" Then
    Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DcboItemName.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If
If CboItemCase.ListIndex = -1 Then
    Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    CboItemCase.SetFocus
    SendKeys "{F4}"
    Exit Sub
End If
If Not TxtSerial Is Nothing Then
    If Trim(Me.TxtSerial.text) <> "" Then
        With Me.Grid
            StrSerial = Trim(Me.TxtSerial.text)
            LngRow = .FindRow(StrSerial, .FixedRows, .ColIndex("Serial"), False, True)
            If LngRow <> -1 Then
                Msg = "Â–« «·”Ì—Ì«· „œŒ· „”»ﬁ« ›Ï «·”ÿ— —ﬁ„ " & LngRow
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End With
    End If
End If
If Val(TxtQuantity.text) = 0 Then
    Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtQuantity.SetFocus
    Exit Sub
End If
If Not IsNumeric(TxtQuantity.text) Then
    Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtQuantity.SetFocus
    Exit Sub
End If
If TxtPrice.text = "" Then
    Msg = "ÌÃ»  ÕœÌœ «·”⁄—"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtPrice.SetFocus
    Exit Sub
End If
If Not IsNumeric(TxtPrice.text) Then
    Msg = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtPrice.SetFocus
    Exit Sub
End If
If Me.TxtQuantity.Tag = "T" Then
    If Not Me.TxtSerial Is Nothing Then
        If Trim(Me.TxtSerial.text) = "" Then
            Msg = "ÌÃ» ﬂ «»… «·”Ì—Ì«· ‰„»—...!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtSerial.SetFocus
            Exit Sub
        End If
    End If
End If
'-------------------------------------------------------------------------------------------------
LngItemID = Val(Me.DcboItemName.BoundText)
'Check if this is Assblied Item
'Or it is a Services Item
RsItems.MoveFirst
RsItems.Find "ItemID=" & LngItemID, , adSearchForward, adBookmarkFirst
If RsItems.BOF Or RsItems.EOF Then
    Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
ElseIf (RsItems("AssbliedItem").Value = True Or _
    RsItems("AssbliedItem").Value = -1 Or RsItems("AssbliedItem").Value = 1) Then
    '’‰› „Ã„⁄
    
ElseIf (RsItems("ItemType").Value = 1) Then
    '’‰› Œœ„…
    
Else
    StrItemDataMsg = "«·’‰› : " & RsItems("ItemName").Value
    StrItemDataMsg = StrItemDataMsg & Chr(13) & "—ﬁ„:" & RsItems("ItemID").Value
    StrItemDataMsg = StrItemDataMsg & Chr(13) & "ﬂÊœ:" & RsItems("ItemCode").Value
    '’‰› ⁄«œÏ
    '-------------------------------------------------------------------------------
    'ÌÃ» «· Õﬁﬁ „‰ «‰ «·ﬂ„Ì…  ”„Õ (›Ï Õ«·… ›Ê« Ì— «·„»Ì⁄«  «Ê „— Ã⁄ «·„‘ —Ì« ) ﬁ»· ≈÷«›… «·’‰›
    'Ê«‰ ≈⁄œ«œ«  «·»—‰«„Ã „⁄œ… ⁄·Ï  ‰Ÿ«„ «·Ã—œ «·›⁄·Ï
    ' ›Ï Õ«·… «‰ ÌﬂÊ‰ „— Ã⁄ „»Ì⁄«  «Ê ›« Ê—… „‘ —Ì«  ›ÌÃ» «· «ﬂœ „‰ „‰⁄  ﬂ—«—
    '«·”Ì—Ì«· ‰„»—
    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Then
        '›« Ê—… »Ì⁄ «Ê „— Ã⁄ „‘ —Ì« 
        If SystemOptions.SysAllowStockNegative = False Then 'Must Check the Stock
            Set RsTemp = New ADODB.Recordset
            Set RsTemp = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, _
                Me.DtpBillDate.Value, , , , Trim(Me.TxtSerial.text))
            If RsTemp.BOF Or RsTemp.EOF Then
                Msg = StrItemDataMsg
                Msg = Msg & "«·ﬂ„Ì… ·«  ”„Õ .. √Ê «‰ «·”Ì—Ì«· «·„œŒ· €Ì— ’ÕÌÕ"
                MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
                Exit Sub
            ElseIf RsTemp("Qty").Value < Val(Me.TxtQuantity.text) Then
                Msg = StrItemDataMsg
                Msg = Msg & Chr(13) & "«·ﬂ„Ì… ·«  ”„Õ .."
                Msg = Msg & Chr(13) & "«·ﬂ„Ì… «·„ÊÃÊœ…:" & RsTemp("Qty").Value
                MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
                Exit Sub
            End If
        End If
    ElseIf Me.GridTrans = ShowPrice Then
        'NO Check Requried
    ElseIf (Me.GridTrans = ReturnSalling Or Me.GridTrans = PurchaseTransaction) And Trim(Me.TxtSerial.text) <> "" Then
        Set RsTemp = New ADODB.Recordset
        Set RsTemp = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, _
            Me.DtpBillDate.Value, , , , Trim(Me.TxtSerial.text))
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Msg = "ÌÊÃœ Õ«·Ì« ›Ï «·„Œ“‰ ‰›” —ﬁ„ «·”Ì—Ì«· ·‰›” «·’‰› ...!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
            Me.TxtSerial.SetFocus
            Exit Sub
        End If
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
        AddItemInGrid .Rows - 1, DcboItemName.BoundText, Val(TxtQuantity.text), _
        Val(Me.TxtPrice.text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.text)
    Else
        AddItemInGrid .Rows - 1, DcboItemName.BoundText, Val(TxtQuantity.text), _
        Val(Me.TxtPrice.text), Me.CboItemCase.ListIndex + 1
    End If
End With
CountItems
If Me.GridTrans <> MoveItems And Me.GridTrans <> OpeningBalance And Me.GridTrans <> ReturnTransaction Then
    Calculate Me.Grid.Rows - 1
End If
If Me.GridTrans = ReturnTransaction Then
    FrmReturnpurchases.Fg_AfterEdit 1, 1
End If
ClearTopDataEntry
'Me.DCboItemCode.SetFocus
'Me.DCboItemCode.SetFocus
Me.Grid.SetFocus
Exit Sub
ErrTrap:
End Sub

Private Sub CmdAddSerialLIst_Click()
Dim LngItemID As Long
Dim LngStoreID As Long
LngItemID = Val(Me.DcboItemName.BoundText)
PopUpSerialList LngItemID
ClearTopDataEntry
End Sub

Private Sub DCboItemCode_Change()
On Error GoTo ErrTrap
Dim StrSQL As String
Dim RsTemp As ADODB.Recordset
Dim LngItemID As Long

If DCboItemCode.BoundText <> "" Then
    DcboItemName.BoundText = DCboItemCode.BoundText
Else
    Exit Sub
End If
LngItemID = Val(DCboItemCode.BoundText)
StrSQL = "select * From TblItems where ItemID=" & DCboItemCode.BoundText
Set RsTemp = New ADODB.Recordset
RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Not (RsTemp.EOF Or RsTemp.BOF) Then
    If Not (IsNull(RsTemp("ItemCase").Value)) = True Then
        If RsTemp("ItemCase").Value = 1 Then
            Me.CboItemCase.ListIndex = 0
        ElseIf RsTemp("ItemCase").Value = 2 Then
            Me.CboItemCase.ListIndex = 1
        End If
    End If
    If RsTemp("HaveSerial").Value = True Then
        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            TxtSerial.Enabled = True
            If Not Me.CmdAddSerialLIst Is Nothing Then
                Me.CmdAddSerialLIst.Enabled = True
            End If
        End If
'            TxtSerial.Tag = "T"
        TxtQuantity.Enabled = False
        TxtQuantity.text = "1"
        TxtQuantity.Tag = "T"
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or _
            Me.GridTrans = Template Then
            TxtPrice.text = GetItemPrice(LngItemID, 1)
        Else
            TxtPrice.text = IIf(IsNull(RsTemp("PurchasePrice").Value), 0, _
            RsTemp("PurchasePrice").Value)
        End If
    ElseIf RsTemp("HaveSerial").Value = False Then
        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            TxtSerial.Enabled = False
            If Not Me.CmdAddSerialLIst Is Nothing Then
                Me.CmdAddSerialLIst.Enabled = False
            End If
        End If
        TxtQuantity.Tag = "F"
        TxtQuantity.Enabled = True
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
            TxtPrice.text = GetItemPrice(LngItemID, 1)
        ElseIf Me.GridTrans = Destruction Then
            '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
            TxtPrice.text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, , , SystemOptions.SysMainStockCostMethod)
        Else
            TxtPrice.text = IIf(IsNull(RsTemp("PurchasePrice").Value), 0, RsTemp("PurchasePrice").Value)
        End If
    End If
End If
Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemCode_GotFocus()
DCboItemCode.BackColor = vbYellow
End Sub

Private Sub DCboItemCode_LostFocus()
DCboItemCode.BackColor = vbWhite
'TxtPrice.BackColor = vbWhite
End Sub

'
Private Sub DcboItemName_Change()
On Error GoTo ErrTrap
Dim StrSQL As String
Dim RsTemp As ADODB.Recordset
Dim LngItemID As Long

If DcboItemName.BoundText <> "" Then
    DCboItemCode.BoundText = DcboItemName.BoundText
Else
    Exit Sub
End If
LngItemID = Val(DCboItemCode.BoundText)
StrSQL = "select * From TblItems where ItemID=" & DcboItemName.BoundText
Set RsTemp = New ADODB.Recordset
RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (RsTemp.EOF Or RsTemp.BOF) Then
    If Not (IsNull(RsTemp("ItemCase").Value)) = True Then
        If RsTemp("ItemCase").Value = 1 Then
            Me.CboItemCase.ListIndex = 0
        ElseIf RsTemp("ItemCase").Value = 2 Then
            Me.CboItemCase.ListIndex = 1
        End If
    End If
    If RsTemp("HaveSerial").Value = True Then
        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            TxtSerial.Enabled = True
            If Not Me.CmdAddSerialLIst Is Nothing Then
                Me.CmdAddSerialLIst.Enabled = True
            End If
        End If
        TxtQuantity.Enabled = False
        TxtQuantity.text = "1"
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
            TxtPrice.text = GetItemPrice(LngItemID, 1)
        Else
            TxtPrice.text = IIf(IsNull(RsTemp("PurchasePrice").Value), 0, RsTemp("PurchasePrice").Value)
        End If
    ElseIf RsTemp("HaveSerial").Value = False Then
        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
            TxtSerial.Enabled = False
            If Not Me.CmdAddSerialLIst Is Nothing Then
                Me.CmdAddSerialLIst.Enabled = False
            End If
        End If
        TxtQuantity.Enabled = True
        TxtQuantity.text = ""
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
            TxtPrice.text = GetItemPrice(LngItemID, 1)
        Else
            TxtPrice.text = IIf(IsNull(RsTemp("PurchasePrice").Value), 0, RsTemp("PurchasePrice").Value)
        End If
    End If
End If
Exit Sub
ErrTrap:
End Sub

'
Private Sub DCboItemName_GotFocus()
DcboItemName.BackColor = vbYellow
If DcboItemName.BoundText = "" Then
    SendKeys "{F4}"
End If
'TxtPrice.BackColor = vbWhite
End Sub

Private Sub DCboItemName_LostFocus()
'DCboItemName.BackColor = vbYellow
DcboItemName.BackColor = vbWhite
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
                Calculate 1, , , True
            End If
        End If
    Case "ClearAll"
        Msg = "”Ê› Ì „ Õ–› ﬂ· «·”ÿÊ— «·„ÊÃÊœ… ›Ï «·ÃœÊ· ."
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ Â–Â «·ŒÿÊ…...ø"
        IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
        If IntRes = vbYes Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
            Me.Grid.Rows = Me.Grid.FixedRows + 1
            Me.Grid.Row = 1
        End If
    Case "FillData"
        If Me.GridTrans = OpeningBalance Then
            If Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Code")) <> "" Then
               Grid.Rows = Grid.Rows + 1
               Me.GridDefaultValue Me.Grid.Rows - 1
            End If
            FrmFillBalance.Show vbModal
        Else
            If Me.Grid.TextMatrix(Me.Grid.Rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
               Me.Grid.Rows = Me.Grid.Rows + 1
               Me.GridDefaultValue Me.Grid.Rows - 1
            End If
            FrmFillItems.DealingForm = Me.GridTrans
            FrmFillItems.Show vbModal
            Calculate 1, , , True
        End If
    Case "Tree"
        If Button.Value = tbrPressed Then
            Set MyAddItemForm = New FrmAddItem
            Set MyAddItemForm.cGrid = Me
            MyAddItemForm.Show
        Else
            CloseMyAddItemForm
        End If
    Case "Barcode"
        If Me.Grid.Parent.XPTxtBillID.text <> "" Then
            Load FrmPrintBarcode
            FrmPrintBarcode.LblID.Caption = Me.Grid.Parent.XPTxtBillID.text
            FrmPrintBarcode.Show vbModal
        End If
    Case "GridCols"
        ShowGridColFrm
End Select
Exit Sub
ErrTrap:
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

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

On Error GoTo ErrTrap
If TxtFillData.text = "T" Then Exit Sub
If Row = -1 Or Col = -1 Then Exit Sub
If RsItems Is Nothing Then Exit Sub

With Me.Grid
    Select Case .ColKey(Col)
        Case "Code"
            If Grid.TextMatrix(Row, Col) <> "" Then
                LngItemID = Val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, adBookmarkFirst
                If RsItems.EOF Or RsItems.BOF Then
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.TextMatrix(Row, Col) = ""
                    Me.Grid.Row = Row
                    Me.Grid.Col = Col
                    Exit Sub
                Else
                    AddItemInGrid Row, LngItemID
                End If
            Else
                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Grid.Cell(flexcpText, Row, Col, Row, .Cols - 1) = ""
                Me.Grid.Row = Row
                Me.Grid.Col = Col
            End If
        Case "Name"
            If Grid.TextMatrix(Row, Col) <> "" Then
                LngItemID = Val(Grid.ComboData)
                RsItems.Find "ItemID=" & LngItemID, , adSearchForward, 1
                If RsItems.EOF Or RsItems.BOF Then
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.TextMatrix(Row, Col) = ""
                    Me.Grid.Row = Row
                    Me.Grid.Col = Col
                    Exit Sub
                Else
                    AddItemInGrid Row, LngItemID
                End If
            Else
                Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & Chr(13)
                'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Grid.TextMatrix(Row, Col) = ""
                Me.Grid.Row = Row
                Me.Grid.Col = Col
                Exit Sub
            End If
        Case "Serial"
            LngItemID = Val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
            SearchResault = Grid.FindRow(Grid.TextMatrix(Row, Grid.ColIndex("Serial")), , Grid.ColIndex("Serial"))
            If SearchResault <> Row And SearchResault > 0 Then
                StrSQL = "Select * From TblItems where ItemID=" & Trim(Grid.TextMatrix(Row, _
                Grid.ColIndex("Code")))
                Set RsSerial = New ADODB.Recordset
                RsSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                StrSQL = "Select * From TblItems where ItemID=" & Trim(Grid.TextMatrix(SearchResault, _
                Grid.ColIndex("Code")))
                Set RsFind = New ADODB.Recordset
                RsFind.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                Msg = " „ «œŒ«· «·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                Msg = Msg + RsSerial("ItemName").Value & Chr(13)
                Msg = Msg + "ﬁ»· –·ﬂ «·Ì «·’‰›" & Chr(13)
                Msg = Msg + RsFind("ItemName").Value & Chr(13)
                Msg = Msg + " √ﬂœ „‰ «·”Ì—Ì«· «·Œ«’ »«·√’‰«›"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                .TextMatrix(Row, Col) = ""
             End If
            '(N)«· √ﬂœ √‰ «·ﬁÿ⁄Â „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ ›Ì «·Õ«·…
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = MoveItems Then
                If SystemOptions.SysAllowStockNegative = False Then
                If TxtModFlag.text = "N" Then
                    If .TextMatrix(Row, .ColIndex("Code")) <> "" Then
                        LngItemID = Val(.TextMatrix(Row, .ColIndex("Code")))
                        StrSQL = "Select * From TblItems where ItemID=" & LngItemID
                        Set RsTemp = New ADODB.Recordset
                        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                        If RsTemp("HaveSerial").Value = True Then
                            If Trim(.TextMatrix(Row, .ColIndex("Serial"))) <> "" Then
                                If StoreName.BoundText <> "" Then
                                    Set RsTest = New ADODB.Recordset
                                    Set RsTest = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, , , True, , Trim(.TextMatrix(Row, .ColIndex("Serial"))))
                                    If RsTest.EOF Or RsTemp.BOF Then
                                        Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«· : "
                                        Msg = Msg + Trim(.TextMatrix(Row, .ColIndex("Serial"))) & Chr(13) & _
                                        "„‰ «·’‰› : " & Trim(.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))) & Chr(13) & _
                                        "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & Row
                                        Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & Chr(13)
                                        Msg = Msg + "Â·  —€»  ›Ì „⁄—›… ”Ì—Ì«· «·ﬁÿ⁄ «·„ÊÃÊœ… ›Ï «·„Œ«“‰"
                                
                                        If (MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + _
                                            vbMsgBoxRtlReading, App.Title)) = vbYes Then
                                            FrmSearchSerial.Tag = RsTemp("ItemCode").Value
                                            If Me.GridTrans = InvoiceTransaction Then
                                                FrmSearchSerial.LngGridRow = Row
                                                FrmSearchSerial.Txt.text = "Search"
                                            ElseIf Me.GridTrans = ReturnTransaction Then
                                                FrmSearchSerial.Txt.text = "Retutn"
                                            ElseIf Me.GridTrans = MoveItems Then
                                                FrmSearchSerial.Txt.text = "Moving"
                                            ElseIf Me.GridTrans = Destruction Then
                                                FrmSearchSerial.Txt.text = "Destruction"
                                            End If
                                            'FrmSearchSerial.Cmd_Click 0
                                            FrmSearchSerial.Show vbModal
                                        End If
                                        'Exit Sub
                                    End If
                                End If
                                RsTest.Close
                                If Me.GridTrans = InvoiceTransaction Then
                                    CalculateInvProfit Row
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
                If Not (IsNumeric(.TextMatrix(Row, .ColIndex("guaranteeTime")))) _
                Or Val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) < 1 _
                Or Val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) > 36 _
                Or Val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) <> _
                Int(Val(.TextMatrix(Row, .ColIndex("guaranteeTime")))) _
                Then
                    Msg = "Â–Â «·ﬁÌ„… €Ì— ’«·Õ…" & Chr(13)
                    Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì… ’ÕÌÕ… »Ì‰ 1 -36 " & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            End If
        Case "Count"
            LngItemID = Val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
            '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
            If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Then
                AddItemPrice LngItemID, Col, Row
            End If
    End Select
End With

If m_GridTrans <> ReturnTransaction And m_GridTrans <> OpeningBalance Then
    If Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 0 Or _
    Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 1 Or _
    Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 Then
         Grid.TextMatrix(Row, Grid.ColIndex("DiscountVal")) = 0
    End If
End If
With Grid
    Select Case .ColKey(Col)
        Case "Count", "Price", "DiscountType", "DiscountVal"
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Then
                Me.Calculate Row, , , True
            ElseIf Me.GridTrans = PurchaseTransaction Then
                Me.Calculate Row
            End If
    End Select
End With
CountItems
HideComment
If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Then
    Me.Calculate Row, , , True
ElseIf Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnTransaction Or _
     Me.GridTrans = Template Or Me.GridTrans = _
    Destruction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = OpeningBalance Then
    Me.Calculate Row
End If
'-------------------------
'Me.Grid.AutoSize 0, Grid.Cols - 1, False
'-------------------------
With Grid
    If .TextMatrix(Row, .ColIndex("Code")) <> "" Then
        EditGrid Row, Col
    Else
        If .Rows = .FixedRows + 1 Then
            .Clear flexClearScrollable, flexClearEverything
            Me.GridDefaultValue 1
        Else
            .RemoveItem Row
        End If
    End If
End With
Exit Sub
ErrTrap:
End Sub
Private Sub EditGrid(Row As Long, Col As Long)
Dim I As Long
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
        For I = Col + 1 To .Cols - 1
            If .ColHidden(I) = False Then
                .Col = I
                Exit For
            End If
        Next I
        .ShowCell .Row, .Col
        GridDefaultValue .Row
        .EditCell
    End If
End With

End Sub

Private Sub Grid_AfterSort(ByVal Col As Long, Order As Integer)
With Grid
    If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
        .Rows = .Rows + 1
    End If
End With
End Sub

Private Sub Grid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
With Me.Grid
    If .ColIndex("AddSerials") <> -1 Then
        .ColWidth(.ColIndex("AddSerials")) = 350
    End If
End With
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrTrap

With Me.Grid
    Select Case .ColKey(Col)
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

If m_GridTrans <> ReturnTransaction And m_GridTrans <> OpeningBalance Then
    If Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 0 Or _
    Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 1 Or _
    Val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 Then
         Grid.TextMatrix(Row, Grid.ColIndex("DiscountVal")) = 0
        If Col = Grid.ColIndex("DiscountVal") Then
            Cancel = True
        End If
    End If
    If Col = Grid.ColIndex("Valu") Then
        Cancel = True
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
Dim TotalNet As Double
Dim DblDiscountTotal As Double
Dim DblRowTotal As Double
On Error GoTo ErrTrap

CountItems
With Grid
    ReSerialGrid
    For RowNum = .FixedRows To .Rows - 1
        .TextMatrix(RowNum, .ColIndex("Valu")) = _
        Val(.TextMatrix(RowNum, .ColIndex("Count"))) * _
        Val(.TextMatrix(RowNum, .ColIndex("Price")))
    Next RowNum
    If .Rows > 1 Then
        Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
    End If
    If Not Me.LblTotalAll Is Nothing Then
        Me.LblTotalAll.Caption = Total
    End If
    DblDiscountTotal = 0
    If .ColIndex("DiscountType") <> -1 Then
        For RowNum = 1 To .Rows - 1
            DblRowTotal = Val(.TextMatrix(RowNum, .ColIndex("Count"))) * Val(.TextMatrix(RowNum, .ColIndex("Price")))
            Select Case Val(.TextMatrix(RowNum, .ColIndex("DiscountType")))
                 Case 0
                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                    DblDiscountTotal = DblDiscountTotal + 0
                Case 1
                    .TextMatrix(RowNum, .ColIndex("Valu")) = DblRowTotal
                    DblDiscountTotal = DblDiscountTotal + 0
                Case 2
                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) - Val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                    DblDiscountTotal = DblDiscountTotal + Val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                Case 3
                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (Val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) / 100))
                    DblDiscountTotal = DblDiscountTotal + ((Val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) * DblRowTotal) / 100)
                Case 4
                    .TextMatrix(RowNum, .ColIndex("Valu")) = 0
                    DblDiscountTotal = DblDiscountTotal + DblRowTotal
            End Select
        Next RowNum
        Total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
        If Me.GridTrans <> ShowPrice And Me.GridTrans <> Template And Me.GridTrans <> Destruction And _
            Me.GridTrans <> ReturnSalling And Me.GridTrans <> ReturnTransaction And Me.GridTrans <> _
            StockSettlement Then
            'Õ”«» ﬁÌ„… «·Œ’„
            If Not CboDiscount_Type Is Nothing Then
                If CboDiscount_Type.ListIndex = 1 Then
                    Total = Total - Val(TxtDiscount_Val.text)
                    DblDiscountTotal = DblDiscountTotal + Val(TxtDiscount_Val.text)
                Else
                    If CboDiscount_Type.ListIndex = 2 Then
                        If TxtDiscount_Val.text <> "" Then
                            DblDiscountTotal = DblDiscountTotal + ((Val(TxtDiscount_Val.text) * Total) / 100)
                            Total = Total * Val(1 - (TxtDiscount_Val.text / 100))
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not Me.LblDiscountsTotal Is Nothing Then
        Me.LblDiscountsTotal.Caption = DisplayCurrency(DblDiscountTotal)
    End If
End With
If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling Then
    CalculateInvProfit LngRow
End If

If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Then
    TotalNet = Total
    '÷—»Ì… Œ’„ Ê≈÷«›…
    If Not Me.TxtAddTax Is Nothing Then
        If Val(Me.TxtAddTax.text) > 0 Then
            Dim DblTemp As Double
            If Not Me.LblTaxAddValue Is Nothing Then
                LblTaxAddValue.Caption = ((Val(Me.TxtAddTax.text) * Total) / 100)
            End If
            TotalNet = Total * (1 - Val((Me.TxtAddTax.text)) / 100)
        End If
    End If
    '÷—»Ì… œ„€…
    If Not Me.TxtStampTax Is Nothing Then
        If Val(Me.TxtStampTax.text) > 0 Then
            If Not Me.LblTaxStampValue Is Nothing Then
                Me.LblTaxStampValue.Caption = Val(Me.TxtStampTax.text)
            End If
            TotalNet = TotalNet - Val(Me.TxtStampTax.text)
        End If
    End If
    '÷—»Ì… «·„»Ì⁄« 
    If TaxCalc = True Then
        If TxtTaxValue.text <> "" Then
            If Val(TxtTaxValue.text) <> 0 Then
                If Not Me.LblTaxSalesValue Is Nothing Then
                    LblTaxSalesValue.Caption = ((Val(Me.TxtTaxValue.text) * Total) / 100)
                End If
                TotalNet = TotalNet + ((Val(Me.TxtTaxValue.text) * Total) / 100)
            End If
        End If
    End If
    '÷—»Ì… Œœ„…
    If Not Me.TxtServiceTax Is Nothing Then
        If Val(Me.TxtServiceTax.text) <> 0 Then
            If Not Me.LblTaxServiceValue Is Nothing Then
                Me.LblTaxServiceValue.Caption = ((Val(Me.TxtServiceTax.text) * Total) / 100)
            End If
            TotalNet = TotalNet + ((Val(Me.TxtServiceTax.text) * Total) / 100)
        End If
    End If
    TotalNet = Format(TotalNet, SystemOptions.SysDefCurrencyForamt)
    TxtTotal.text = TotalNet
Else
    TxtTotal.text = Total
End If

If Not BolSavae = False Then
    If TxtModFlag.text = "N" Or TxtModFlag.text = "E" Then
        If Val(TxtValueCash.text) + Val(TxtValueDelay.text) + _
            Val(TxtValuechque.text) > Val(Total) Then
            Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & Chr(13)
            Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Calculate = False
            Exit Function
        End If
        If Val(TxtValueCash.text) + Val(TxtValueDelay.text) + Val(TxtValuechque.text) < Val(Total) Then
            Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & Chr(13)
            Msg = Msg + "√ﬁ· „‰ ≈Ã„«·Ì «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Calculate = False
            Exit Function
        End If
    End If
End If
Calculate = True

Exit Function
ErrTrap:
End Function
Public Property Get GridTrans() As GridTransType
GridTrans = m_GridTrans
End Property
Public Property Let GridTrans(ByVal vNewValue As GridTransType)
If (vNewValue <> InvoiceTransaction And vNewValue <> PurchaseTransaction And vNewValue <> ReturnTransaction And vNewValue <> GridTransType.ShowPrice And vNewValue <> Template And vNewValue <> Destruction And vNewValue <> ReturnSalling And vNewValue <> MoveItems And vNewValue <> OpeningBalance And vNewValue <> StockSettlement) Then
Else
    m_GridTrans = vNewValue
End If
End Property
Public Sub GridDefaultValue(Crow As Long)
On Error GoTo ErrTrap
With Grid
    .TextMatrix(Crow, .ColIndex("ItemCase")) = 1
    If .TextMatrix(Crow, .ColIndex("Count")) = "" Then
        .TextMatrix(Crow, .ColIndex("Count")) = 1
    End If
    
    If .ColIndex("DiscountType") <> -1 Then
        .TextMatrix(Crow, .ColIndex("DiscountType")) = 1
    End If
   
End With
Exit Sub
ErrTrap:
End Sub
Public Sub FillGrid()

Dim StrList As String
Dim RsNote As ADODB.Recordset
Dim StrSQL As String
Dim Rs As ADODB.Recordset
On Error GoTo ErrTrap

m_ShowQtyToolTip = True
With Grid
    If SystemOptions.UserInterface = EnglishInterface Then
        .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
        .TextMatrix(0, .ColIndex("Code")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = "Item Name"
        .TextMatrix(0, .ColIndex("Serial")) = "Item Serial"
        .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "Have Serial"
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
    If m_GridTrans <> OpeningBalance Then
        If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("DiscountType")) = "#1;·«ÌÊÃœ Œ’„|#2;Œ’„ »ﬁÌ„…|#3;Œ’„ »‰”»…|#4;„Ã«‰Ì"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .ColComboList(.ColIndex("DiscountType")) = "#1;NO Discount|#2;Value Discount|#3;Percent Discount|#4;Free"
        End If
    End If
    
    .ColAlignment(.ColIndex("Serial")) = flexAlignLeftCenter
    '--------------------------------------------
    .Cols = .Cols + 1
    .ColWidthMin = 350
    .ColKey(.Cols - 1) = "AddSerials"
    .ColComboList(.ColIndex("AddSerials")) = "..."
    
    .ColPosition(.ColIndex("AddSerials")) = 5 ' .ColPosition(.ColIndex("Serail")/
    
    .Cols = .Cols + 1
    .ColKey(.Cols - 1) = "Remarks"
    .ColComboList(.ColIndex("Remarks")) = "..."
    '.ColPosition(.ColIndex("Remarks")) = .Cols - 1
    If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(0, .ColIndex("Remarks")) = "„·«ÕŸ« "
    Else
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End If
    '--------------------------------------------
    '≈÷«›… ⁄„Êœ ‰Ê⁄ «·’‰› ≈› —«÷Ì« ›Ï ﬂ· «·‘«‘« 
    If .ColIndex("ItemType") = -1 Then
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "ItemType"
'        .ColHidden(.ColIndex("ItemType")) = True
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
    '--------------------------------------------
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
    '--------------------------------------------
        
.Cols = .Cols + 1
    .ColKey(.Cols - 1) = "UnitID"
    .ColPosition(.ColIndex("UnitID")) = .Cols - 1
    .TextMatrix(0, .ColIndex("UnitID")) = "ÊÕœ… «·’‰›"
    .ColAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
    .FixedAlignment(.ColIndex("UnitID")) = flexAlignRightCenter
    
    'Set the Grid Drag Drop Cap
    .OLEDropMode = flexOLEDropAutomatic
    .OLEDragMode = flexOLEDragAutomatic
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
    StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Update TblItems Set ItemType=0 Where ItemType IS NULL"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Update TblItems Set AssbliedItem=0 Where AssbliedItem IS NULL"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Update TblItems Set RelatedItem=0 Where RelatedItem IS NULL"
    Cn.Execute StrSQL, , adExecuteNoRecords
End If
Set RsItems = New ADODB.Recordset
StrSQL = "Select * From TblItems "

If Me.GridTrans = Destruction Then
    StrSQL = StrSQL
GoTo order_By
End If



If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Then
    StrSQL = StrSQL + " Where (TblItems.AssbliedItem=1 OR TblItems.AssbliedItem=0)"
Else
    StrSQL = StrSQL + " Where (TblItems.AssbliedItem=0)"
End If

If Me.GridTrans = InvoiceTransaction Then
    StrSQL = StrSQL + " AND (TblItems.ItemType=0 OR TblItems.ItemType=1)"
Else
    StrSQL = StrSQL + " AND (TblItems.ItemType=0)"
End If

order_By:

StrSQL = StrSQL + " Order By ItemID"


RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (RsItems.BOF Or RsItems.EOF) Then
    StrList = Me.Grid.BuildComboList(RsItems, "ItemName", "ItemID")
    If StrList <> "" Then
        Grid.ColComboList(Grid.ColIndex("Name")) = "|" & StrList
    End If
    StrList = Grid.BuildComboList(RsItems, "ItemCode", "ItemID")
    If StrList <> "" Then
        Grid.ColComboList(Grid.ColIndex("Code")) = "|" & StrList
    End If
    RsItems.MoveFirst
End If
'--------------------------------------
FillItemData
DCboItemCode.TabIndex = Me.Grid.TabIndex + 1
DcboItemName.TabIndex = Me.DCboItemCode.TabIndex + 1
CboItemCase.TabIndex = DcboItemName.TabIndex + 1
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
    LblItemsCount.BackColor = &H404040
    LblItemsCount.ForeColor = &HFFFF&
    LblItemsCount.Alignment = vbCenter
    LblItemsCount.RightToLeft = True
End If
GetMySetting

Exit Sub
ErrTrap:
'Stop
'Resume
End Sub

Private Sub Grid_BeforeSort(ByVal Col As Long, Order As Integer)
Dim I As Integer
With Grid
    If .TextMatrix(.Rows - 1, .ColIndex("Code")) = "" Then
        .RemoveItem .Rows - 1
         Calculate 1
    End If
End With

End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim LngItemID As Long
Dim LngStoreID As Long
Dim Frm As FrmGridAddItemComment
On Error GoTo ErrTrap
With Me.Grid
    Select Case .ColKey(Col)
        Case "NewItem"
            On Error GoTo ErrTrap
            If Me.TxtModFlag.text = "N" Or Me.TxtModFlag.text = "E" Then
                FrmAddNewItem.DealingForm = Me.GridTrans
                FrmAddNewItem.Show vbModal
            End If
        Case "AddSerials"
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Or Me.GridTrans = MoveItems Then
                If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexUnchecked Then
                    Exit Sub
                End If
                LngItemID = Val(.TextMatrix(Row, .ColIndex("Name")))
                LngStoreID = Val(Me.StoreName.BoundText)
                If LngItemID = 0 Or LngStoreID = 0 Then
                    Exit Sub
                End If
                PopUpSerialList LngItemID, Row, Me.Grid.ColIndex("Serial")
            End If
        Case "Remarks"
            ItemComment Row, Col
        End Select
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
Public Sub FillItemData()
On Error GoTo ErrTrap
Dim StrSQL As String
Dim Rs As ADODB.Recordset
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

If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Then
    Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1
    Dcombos.GetItemsNames Me.DcboItemName, -1, -1
ElseIf Me.GridTrans = PurchaseTransaction Then
    Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0
    Dcombos.GetItemsNames Me.DcboItemName, 0, 0
Else
    Dcombos.GetItemsCodes Me.DCboItemCode, 0, 0
    Dcombos.GetItemsNames Me.DcboItemName, 0, 0
End If

If Me.GridTrans = InvoiceTransaction Then
    StrSQL = StrSQL + " AND TblItems.ItemType=0 OR TblItems.ItemType=1"
Else
    StrSQL = StrSQL + " AND TblItems.ItemType=0"
End If
Set cSearchDcbo(0) = New clsDCboSearch
Set cSearchDcbo(0).Client = Me.DcboItemName

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
If Me.TxtModFlag.text = "R" Then
    With Me.Grid
        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Col = .ColIndex("Serial") Then
            If Trim(.TextMatrix(.Row, .ColIndex("Serial"))) = "" Then
                Exit Sub
            End If
            Load FrmSerialData
            With FrmSerialData
                .XPTxtCode.text = Trim(Grid.TextMatrix(Grid.Row, Grid.ColIndex("Serial")))
                .DcboItemName.BoundText = Me.Grid.TextMatrix(Grid.Row, Grid.ColIndex("Name"))
                .Show
                .Cmd_Click 0
                .ZOrder 0
            End With
        ElseIf .Col = .ColIndex("Code") Or .Col = .ColIndex("Name") Or .Col = .ColIndex("ItemCase") Then
            If Val(.TextMatrix(.Row, .ColIndex("Name"))) <> 0 Then
                Load FrmSelectData
                FrmSelectData.DcboItemName.BoundText = Val(.TextMatrix(.Row, .ColIndex("Name")))
                FrmSelectData.TxtItemCode.text = Val(.TextMatrix(.Row, .ColIndex("Name")))
                If Not Me.StoreName Is Nothing Then
                    FrmSelectData.DcboStores.BoundText = Me.StoreName.BoundText
                End If
                FrmSelectData.Show
            End If
        ElseIf .Col = .ColIndex("Remarks") And Me.TxtModFlag.text = "R" Then
            ItemComment Grid.Row, Grid.Col
        End If
    End With
End If
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LngMouseRow As Long
Dim StrItemSerial As String
Dim LngStoreID As Long

If Button = vbRightButton Then
    With Me.Grid
        LngMouseRow = .MouseRow
        If LngMouseRow = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If Val(.TextMatrix(LngMouseRow, .ColIndex("Name"))) <> 0 Then
            LngCurrentItemID = Val(.TextMatrix(LngMouseRow, .ColIndex("Name")))
            If Me.Grid.Cell(flexcpChecked, LngMouseRow, .ColIndex("HaveSerial")) = flexChecked Then
                If .TextMatrix(LngMouseRow, .ColIndex("Serial")) <> "..." Or _
                    .TextMatrix(LngMouseRow, .ColIndex("Serial")) <> "" Then
                    StrItemSerial = Trim$(.TextMatrix(LngMouseRow, .ColIndex("Serial")))
                End If
            Else
                StrItemSerial = ""
            End If
            If Not Me.StoreName Is Nothing Then
                LngStoreID = Val(Me.StoreName.BoundText)
            End If
            SetItemReportsMenu LngCurrentItemID, Me.Grid.Parent, StrItemSerial, LngStoreID
        End If
    End With
End If
End Sub

Private Sub Grid_OLEDragDrop(Data As VSFlex8UCtl.VSDataObject, _
Effect As Long, ByVal Button As Integer, _
ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim StrTempDropID As String
Dim LngRow As Long
On Local Error GoTo ErrTrap
If Button = 0 Then ' vbLeftButton
    StrTempDropID = Data.GetData(vbCFText)
    With Me.Grid
        If StrTempDropID = "" Then
            Effect = vbDropEffectNone
        Else
            LngRow = .MouseRow
            If LngRow = -1 Or LngRow = 0 Then 'test if  any row under mouse curser
                LngRow = NewEmptyRow
            End If
            Effect = vbDropEffectNone
            If LngRow = -1 Then Exit Sub
            AddItemInGrid LngRow, Val(StrTempDropID)
            'Unload FrmItemTip
        End If
    End With
End If
Exit Sub
ErrTrap:
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Rs As ADODB.Recordset
Dim StrSQL As String
Dim StrList As String
Dim LngItemID As Long
With Me.Grid
    If Me.GridTrans = InvoiceTransaction Then
        Select Case .ColKey(Col)
            Case "Price"
                If SystemOptions.UserInvoiceChangePrice = 0 Then
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
            LngItemID = Val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))
            If LngItemID = 0 Then
                Cancel = True
            Else
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & _
                "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                Set Rs = New ADODB.Recordset
                Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If Not (Rs.BOF Or Rs.EOF) Then
                    StrList = .BuildComboList(Rs, "UnitName", "UnitID")
                    Grid.ColComboList(.ColIndex("UnitID")) = "|" & StrList
                Else
                     Cancel = True
                End If
            End If

        End Select
    End If
End With

End Sub

Private Sub LblDiscountsTotal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblDiscountsTotal.ToolTipText = WriteNo(LblDiscountsTotal.Caption, 0, True)
End Sub

Private Sub LblTotalAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblTotalAll.ToolTipText = WriteNo(LblTotalAll.Caption, 0, True)
End Sub


Private Sub m_FrmItemsNewPrices_Unload(Cancel As Integer)
Set m_FrmItemsNewPrices = Nothing
End Sub


Private Sub TxtAddTax_Change()
If Me.TxtModFlag.text = "E" Or Me.TxtModFlag.text = "N" Then
    Me.Calculate 1, , , True
End If
End Sub

Private Sub TxtAddTax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAddTax.text, 0)
End Sub


Private Sub TxtModFlag_Change()
Dim I  As Integer

If Me.TxtModFlag.text = "R" Then
    Me.Grid.OLEDropMode = flexOLEDropNone
    Me.Grid.Editable = flexEDNone
    If Not Me.GrdTBar Is Nothing Then
        For I = 1 To Me.GrdTBar.Buttons.count
            Me.GrdTBar.Buttons(I).Enabled = False
        Next I
        Me.GrdTBar.Buttons("Barcode").Enabled = True
        'Me.GrdTBar.Buttons("Options").Enabled = True
        Me.GrdTBar.Buttons("GridCols").Enabled = True
    End If
ElseIf Me.TxtModFlag.text = "E" Then
    Me.Grid.OLEDropMode = flexOLEDropAutomatic
    Me.Grid.Editable = flexEDKbdMouse
    If Not Me.GrdTBar Is Nothing Then
        For I = 1 To Me.GrdTBar.Buttons.count
            Me.GrdTBar.Buttons(I).Enabled = True
        Next I
        Me.GrdTBar.Buttons("Barcode").Enabled = False
      '  Me.GrdTBar.Buttons("Options").Enabled = True
        Me.GrdTBar.Buttons("GridCols").Enabled = True
    End If
ElseIf Me.TxtModFlag.text = "N" Then
    If Not Me.LblTotalAll Is Nothing Then
        Me.LblTotalAll.Caption = 0
    End If
    If Not Me.TxtTotal Is Nothing Then
        Me.TxtTotal.text = 0
    End If
    Me.Grid.OLEDropMode = flexOLEDropAutomatic
    Me.Grid.Editable = flexEDKbdMouse
    If Not Me.GrdTBar Is Nothing Then
        For I = 1 To Me.GrdTBar.Buttons.count
            Me.GrdTBar.Buttons(I).Enabled = True
        Next I
        Me.GrdTBar.Buttons("Barcode").Enabled = False
        'Me.GrdTBar.Buttons("Options").Enabled = True
        Me.GrdTBar.Buttons("GridCols").Enabled = True
    End If
End If
    
End Sub
Private Sub GetMySetting()

Dim StrSetting As String
Dim StrShowSet As String
Dim FrmName As String

Dim VarCols As Variant
Dim VarColSet As Variant
Dim I As Integer
On Error Resume Next
FrmName = Me.Grid.Parent.name
StrSetting = GetSetting(SystemOptions.SysRegsAppPath, "Interface Setting" & "\" & User_ID, FrmName, "")
If StrSetting = "" Then
    Exit Sub
End If

VarCols = Split(StrSetting, ";", , vbTextCompare)
If UBound(VarCols) > 0 Then
    For I = 0 To UBound(VarCols)
        VarColSet = Empty
        VarColSet = Split(CStr(VarCols(I)), "-", , vbTextCompare)
        With Grid
            .ColPosition(.ColIndex(CStr(VarColSet(0)))) = CLng(VarColSet(1))
        End With
    Next I
End If
StrShowSet = GetSetting(SystemOptions.SysRegsAppPath, "Cols Setting" & "\" & User_ID, FrmName, "")
If StrShowSet = "" Then
    Exit Sub
End If
VarCols = Split(StrShowSet, ";", , vbTextCompare)
If UBound(VarCols) > 0 Then
    For I = 0 To UBound(VarCols)
        VarColSet = Empty
        VarColSet = Split(CStr(VarCols(I)), "-", , vbTextCompare)
        With Grid
            .ColHidden(.ColIndex(CStr(VarColSet(0)))) = CBool(VarColSet(1))
        End With
    Next I
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
Dim I As Integer

Load FrmGridColsShow
With FrmGridColsShow.FG
    .Rows = .FixedRows
    '-----------------
        If Me.Grid.ColIndex("Ser") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„”·”· «·’‰›"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Ser"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Ser"))
        End If
        '-----------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·’‰›"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Code"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Code"))
        '------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
         .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«”„ «·’‰›"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Name"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Name"))
        '------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
         .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemCase"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCase"))
        '-------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "·Â ”Ì—Ì«·"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "HaveSerial"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("HaveSerial"))
        '---------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "≈÷«›… «·”Ì—Ì«·"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "AddSerials"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AddSerials"))
        '-------------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·”Ì—Ì«·"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Serial"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Serial"))
        '-------------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì…"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Count"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Count"))
        '-------------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·”⁄—"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Price"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Price"))
        '-------------------
        If Me.Grid.ColIndex("guaranteeTime") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„œ… «·÷„«‰ »«·‘Â—"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "guaranteeTime"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("guaranteeTime"))
        End If
        '-------------------
        If Me.Grid.ColIndex("DiscountType") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·Œ’„"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DiscountType"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountType"))
        End If
        '-------------------
        If Me.Grid.ColIndex("DiscountVal") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·Œ’„"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "DiscountVal"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountVal"))
        End If
        '-------------------
        .Rows = .Rows + 1
        I = I + 1
        .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
        .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "«·ﬁÌ„…"
        .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "Valu"
        .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Valu"))
        '-------------------
        '«··Ê‰
        If Me.Grid.ColIndex("ColorID") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "·Ê‰ «·’‰›"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ColorID"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ColorID"))
        End If
        '---------------------
        '«··Ê‰
        If Me.Grid.ColIndex("ItemSize") <> -1 Then
            .Rows = .Rows + 1
            I = I + 1
            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = I
            .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "„ﬁ«” «·’‰›"
            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ItemSize"
            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemSize"))
        End If
        '---------------------
    .AutoSize 0, .Cols - 1, False
End With
Set FrmGridColsShow.FgGrid = Me.Grid
FrmGridColsShow.Show vbModal
End Sub
Public Sub SaveMySetting()
Dim I As Integer
Dim StrTemp As String
Dim StrShow As String
Dim FrmName As String
FrmName = Me.Grid.Parent.name

For I = 0 To Grid.Cols - 1
    StrTemp = StrTemp & Grid.ColKey(I) & "-" & I & ";"
    StrShow = StrShow & Grid.ColKey(I) & "-" & Grid.ColHidden(I) & ";"
Next I
StrTemp = Trim(StrTemp)
StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
StrShow = Trim(StrShow)
StrShow = Mid(StrShow, 1, Len(StrShow) - 1)
SaveSetting SystemOptions.SysRegsAppPath, "Interface Setting" & "\" & User_ID, FrmName, StrTemp
SaveSetting SystemOptions.SysRegsAppPath, "Cols Setting" & "\" & User_ID, FrmName, StrShow

'-----------------------------------------

End Sub

Private Sub TxtPrice_GotFocus()
TxtPrice.BackColor = vbYellow
SelectText TxtPrice
If Me.DcboItemName.BoundText <> "" Then
    If m_ShowQtyToolTip = True Then
        ShowItemTip ItemPriceToolTip, Val(Me.DcboItemName.BoundText), _
        0, 0, , False, Me.TxtPrice.hwnd
    End If
End If
End Sub

Private Sub TxtPrice_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    CmdAddData_Click
    'CmdAddData.SetFocus
    'DCboItemCode.SetFocus
End If
Exit Sub
ErrTrap:
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPrice.text, 0)
End Sub

Private Sub TxtPrice_LostFocus()
'TxtPrice.BackColor = vbYellow
TxtPrice.BackColor = vbWhite
HideComment
End Sub

Private Sub TxtQuantity_GotFocus()
TxtQuantity.BackColor = vbYellow
End Sub


Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.text, 0)
End Sub


Private Sub TxtQuantity_LostFocus()

TxtQuantity.BackColor = vbWhite
End Sub
Private Sub AddItemPrice(LngItemID As Long, Col As Long, Row As Long)
With Grid
    If TxtModFlag.text = "N" Or TxtModFlag.text = "E" Then
        If Col = .ColIndex("Count") Or Col = .ColIndex("Code") Or Col = .ColIndex("Name") Then
            .TextMatrix(Row, .ColIndex("Price")) = _
            GetItemPrice(LngItemID, Val(.TextMatrix(Row, .ColIndex("Count"))))
        End If
    End If
End With
End Sub

Private Function CalculateInvProfit(LngRow As Long) As Double
Dim DblItemProfit As Double
Dim DblItemCost As Double
Dim DblTotalItemCost As Double
Dim LngItemID As Long
Dim IntHaveSerial As Integer
Dim DblInvProfit As Double
Dim StrItemSerial As String
Dim StrTransID As String
On Error GoTo ErrTrap
If Me.TxtModFlag.text <> "N" And Me.TxtModFlag.text <> "E" Then
    Exit Function
End If
If Me.Grid.Rows = Me.Grid.FixedRows Then Exit Function
With Me.Grid
    LngItemID = Val(Me.Grid.TextMatrix(LngRow, Grid.ColIndex("Code")))
    If LngItemID <> 0 Then
        Select Case Me.Grid.Cell(flexcpChecked, LngRow, Me.Grid.ColIndex("HaveSerial"))
            Case flexChecked
                IntHaveSerial = 1
            Case flexUnchecked
                IntHaveSerial = 2
        End Select
        StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))
        If Not (IntHaveSerial = 1 And StrItemSerial = "") Then
            If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
                DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID, ModernWeightAverage)
            Else
                DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID)
            End If
            
            Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemCostPrice")) = DblItemCost
            
            If Me.Grid.ColIndex("ItemProfit") = -1 Then Exit Function
            DblTotalItemCost = (DblItemCost * _
                Val(Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("Count"))))
                
            Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemProfit")) = _
            Val(.TextMatrix(LngRow, .ColIndex("Valu"))) - DblTotalItemCost
            
            '------------------------------------------------------------------
            If Val(.TextMatrix(LngRow, .ColIndex("Valu"))) = DblTotalItemCost Then
                Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbYellow
            ElseIf Val(.TextMatrix(LngRow, .ColIndex("Valu"))) < DblTotalItemCost Then
                Me.Grid.Cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbRed
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
Exit Function
ErrTrap:
End Function
Public Sub AddItemInGrid(LngRow As Long, LngItemID As Long, _
Optional DblQty As Double = 1, _
Optional SngItemPrice As Single = 0, _
Optional IntItemCase As Integer = 0, _
Optional StrItemSerial As String = "")

Dim RsParts As ADODB.Recordset
Dim StrSQL  As String
Dim I As Long, j As Long
Dim DblPartQty As Double
Dim LngPartItemIndex As Long
Dim LngNewPartRow As Long
Dim LngAddCount As Long
Dim StrSerial As String
Dim Frm As FrmAddItemAttachs
Dim LngAttachedItemID As Long
Dim DblAttachedQty As Double
Dim DblAttachedPrice As Single

If LngRow = 0 Then
    LngRow = Me.NewEmptyRow
ElseIf LngRow > Me.Grid.Rows - 1 Then
    Me.Grid.Rows = LngRow + 1
End If

If Not (RsItems.BOF Or RsItems.EOF) Then
    If RsItems("ItemID").Value <> LngItemID Then
        RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
    End If
End If

If Not (RsItems.BOF Or RsItems.EOF) Then
    With Me.Grid
        If RsItems("AssbliedItem").Value = True Then
            StrSQL = "SELECT TableID, ItemID, PartItemID, PartItemQty, PartItemPrice"
            StrSQL = StrSQL + " FROM TblItemsParts Where ItemID=" & LngItemID
            StrSQL = StrSQL + " Order By TableID"
            Set RsParts = New ADODB.Recordset
            RsParts.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RsParts.EOF Or RsParts.BOF) Then
                For I = 0 To RsParts.RecordCount - 1
                    LngPartItemIndex = LngPartItemIndex + 1
                    '--------------------------------------
                    'Add to grid the New Parts Rows
                    '--------------------------------------
                    DblPartQty = IIf(IsNull(RsParts("PartItemQty").Value), 0, RsParts("PartItemQty").Value)
                    DblPartQty = DblPartQty * DblQty
                    LngNewPartRow = LngRow + (LngPartItemIndex - 1)
                    Me.Grid.AddItem "", LngNewPartRow
                    RsItems.MoveFirst
                    AddItemInGrid LngNewPartRow, RsParts("PartItemID").Value, DblPartQty
                    RsParts.MoveNext
                Next I
            End If
            
            '--------≈÷«›… «·√’‰«› «·„·Õﬁ… »«·’‰›
            If Not (RsItems.BOF Or RsItems.EOF) Then
                If RsItems("ItemID").Value <> LngItemID Then
                    RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
                End If
            End If
            If RsItems("RelatedItem").Value = True Then
                Set Frm = New FrmAddItemAttachs
                Frm.LoadItemData LngItemID
                Frm.Show vbModal
                If Frm.UserCancelled = True Then
                    Unload Frm
                Else
                    For j = Frm.FG.FixedRows To Frm.FG.Rows - 1
                        If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                            LngAttachedItemID = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                            DblAttachedQty = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                            DblAttachedPrice = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                            AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice
                        End If
                    Next j
                    Unload Frm
                End If
            End If
            Exit Sub
        End If
        .TextMatrix(LngRow, .ColIndex("Code")) = LngItemID
        .TextMatrix(LngRow, .ColIndex("Name")) = LngItemID
        .TextMatrix(LngRow, .ColIndex("HaveSerial")) = RsItems("HaveSerial").Value
        If .ColIndex("guaranteeTime") <> -1 Then
            .TextMatrix(LngRow, .ColIndex("guaranteeTime")) = _
                IIf(IsNull(RsItems("GuaranteeValue").Value), 0, RsItems("GuaranteeValue").Value)
        End If
        'DblQty = Val(.TextMatrix(LngRow, .ColIndex("Count")))
        DblQty = IIf(DblQty = 0, 1, DblQty)
        .TextMatrix(LngRow, .ColIndex("Count")) = DblQty
        .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1 'set the defalut item case to new case
        If .ColIndex("DiscountType") <> -1 Then
            .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
        End If
        If SngItemPrice = 0 Then
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or _
                Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Then
                'Put Sales Price
                .TextMatrix(LngRow, .ColIndex("Price")) = GetItemPrice(LngItemID, 1)
            ElseIf Me.GridTrans = Destruction Then
                '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                StrSerial = Trim$(.TextMatrix(LngRow, .ColIndex("Serial")))
                .TextMatrix(LngRow, .ColIndex("Price")) = _
                    ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod)
            Else
                'Put Purchase Price
                .TextMatrix(LngRow, .ColIndex("Price")) = IIf(IsNull(RsItems("PurchasePrice").Value), 0, RsItems("PurchasePrice").Value)
            End If
        Else
            'Put the Sended Price
            .TextMatrix(LngRow, .ColIndex("Price")) = SngItemPrice
        End If
        '----------------------------------
        If IntItemCase = 0 Then
            If IsNull(RsItems("ItemCase").Value) Then
                .TextMatrix(LngRow, .ColIndex("ItemCase")) = 1
            Else
                .TextMatrix(LngRow, .ColIndex("ItemCase")) = RsItems("ItemCase").Value
            End If
        Else
            .TextMatrix(LngRow, .ColIndex("ItemCase")) = IntItemCase
        End If
        .TextMatrix(LngRow, .ColIndex("Serial")) = StrItemSerial
'        If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
'            .TextMatrix(LngRow, .ColIndex("Serial")) = TxtSerial.text
'        End If
        
        If Me.GridTrans <> OpeningBalance And Me.GridTrans <> ReturnTransaction Then
            .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
            .TextMatrix(LngRow, .ColIndex("DiscountVal")) = 0
        End If
        If Me.Grid.ColIndex("ItemType") <> -1 Then
            If Not (IsNull(RsItems("ItemType").Value)) Then
                .TextMatrix(LngRow, .ColIndex("ItemType")) = RsItems("ItemType").Value
                If Val(.TextMatrix(LngRow, .ColIndex("ItemType"))) = 1 Then
                    .Cell(flexcpPicture, LngRow, .ColIndex("Ser")) = _
                        MDIFrmMain.ImgLstTree.ListImages("Maintenance").Picture
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
        .TextMatrix(LngRow, .ColIndex("ColorID")) = 1
        .TextMatrix(LngRow, .ColIndex("ItemSize")) = ""
        '-----------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
        Else
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignLeftCenter
        End If
        '----------------------------------
        '--------≈÷«›… «·√’‰«› «·„·Õﬁ… »«·’‰›
        If RsItems("RelatedItem").Value = True Then
            Set Frm = New FrmAddItemAttachs
            Frm.LoadItemData LngItemID
            Frm.Show vbModal
            If Frm.UserCancelled = True Then
                Unload Frm
            Else
                For j = Frm.FG.FixedRows To Frm.FG.Rows - 1
                    If Frm.FG.Cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                        LngAttachedItemID = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                        DblAttachedQty = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                        DblAttachedPrice = Val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                        AddItemInGrid 0, LngAttachedItemID, DblAttachedQty, DblAttachedPrice
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
  Dim RsUnitData As New ADODB.Recordset
 Set RsUnitData = New ADODB.Recordset
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & _
        "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & _
        "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & _
        "TblItemsUnits.FactorBySmallUnit "
        StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & _
        "TblUnites.UnitID"
        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
        StrSQL = StrSQL + " AND DefaultUnit=1"
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            .Cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").Value
            .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").Value
        End If
        RsUnitData.Close
        Set RsUnitData = Nothing
        PutDefCount LngRow
       
   '====================================================================

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
Dim I As Long
Dim Rs As ADODB.Recordset
Dim StrSQL As String
Dim LngTransID As Long


If Me.GridTrans <> InvoiceTransaction Then Exit Sub
Load FrmInvProfit
FrmInvProfit.CostMethod = SystemOptions.SysMainStockCostMethod

StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & _
"TransactionTypes.TransactionTypeName "
StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN Transactions ON " & _
"TransactionTypes.Transaction_Type = Transactions.Transaction_Type "

Set Rs = New ADODB.Recordset
Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
FrmInvProfit.FgItems.Rows = Me.Grid.Rows
FrmInvProfit.FgItems.FixedRows = Me.Grid.FixedRows

For I = Me.Grid.FixedRows To Me.Grid.Rows - 1
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("ItemID")) = _
    Me.Grid.Cell(flexcpText, I, Me.Grid.ColIndex("Code"))
    
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("ItemCode")) = _
    Me.Grid.Cell(flexcpTextDisplay, I, Me.Grid.ColIndex("Code"))
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("ItemName")) = _
    Me.Grid.Cell(flexcpTextDisplay, I, Me.Grid.ColIndex("Name"))
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("CostPrice")) = _
    Me.Grid.TextMatrix(I, Me.Grid.ColIndex("ItemCostPrice"))
    If SystemOptions.SysMainStockCostMethod = LastPurPriceType Then
        LngTransID = Val(Me.Grid.TextMatrix(I, Me.Grid.ColIndex("PofTransID")))
        If LngTransID <> 0 Then
            If Not (Rs.BOF Or Rs.EOF) Then
                Rs.Find "Transaction_ID=" & LngTransID & "", , adSearchForward, 1
                If Not (Rs.BOF Or Rs.EOF) Then
                    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("TransSerial")) = _
                        IIf(IsNull(Rs("Transaction_Serial").Value), "", Rs("Transaction_Serial").Value)
                    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("TransType")) = _
                        IIf(IsNull(Rs("TransactionTypeName").Value), "", Rs("TransactionTypeName").Value)
                Else
                    Rs.MoveFirst
                End If
            End If
        End If
    End If
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("InvQty")) = _
    Me.Grid.TextMatrix(I, Me.Grid.ColIndex("Count"))
    
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("InvQtyCost")) = _
    Val(FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("InvQty"))) * _
    Val(FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("CostPrice")))
    
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("ItemSalePrice")) = _
    Me.Grid.TextMatrix(I, Me.Grid.ColIndex("Price"))
    'DiscountVal
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("Discount")) = _
    Me.Grid.TextMatrix(I, Me.Grid.ColIndex("DiscountVal"))
    
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("ItemProfit")) = _
    Me.Grid.TextMatrix(I, Me.Grid.ColIndex("ItemProfit"))
    
    FrmInvProfit.FgItems.TextMatrix(I, FrmInvProfit.FgItems.ColIndex("CalculateType")) = ""
    
Next I
    

With FrmInvProfit.FgItems
    .Cell(flexcpBackColor, .FixedRows, .ColIndex("TransType"), .Rows - 1, .ColIndex("TransType")) = &H808000
    .Cell(flexcpBackColor, .FixedRows, .ColIndex("TransSerial"), .Rows - 1, .ColIndex("TransSerial")) = &H808000
    .AutoSize 0, .Cols - 1, False
    FrmInvProfit.lblInvProf.Caption = _
        .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemProfit"), .Rows - 1, .ColIndex("ItemProfit"))
End With

FrmInvProfit.Show vbModeless, MDIFrmMain
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
Dim I As Long
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

Private Sub ShowItemTip(ToolTipType As ItemRowToolTip, StrItemID As String, _
    LngShowRow As Long, LngShowCol As Long, _
    Optional bSetFocus As Boolean = False, _
    Optional BolShowInGrid As Boolean = True, _
    Optional LngHwnd As Long)
    
'------------------------------

'------------------------------

Dim uPoint As POINTAPI
Static lNoteRow As Long, lNoteCol As Long, r As Long, C As Long
Dim lLeft  As Single
Dim LTop As Single
Dim StrSQL As String
Dim Rs As ADODB.Recordset
Dim Acmd As ADODB.Command
Dim XParItemID As ADODB.Parameter
Dim XParTransType As ADODB.Parameter
Dim I As Integer
Dim LngItemID As Long
Dim ItemTransInfo As LastItemTransInfo
Dim xRect As RECT

HideComment
LngItemID = Val(StrItemID)
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
            .LblData(0).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Item_Code"))
            .LblData(1).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Item_Name"))
            .LblData(2).Caption = Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Group_Name"))
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
            LngItemID = Val(StrItemID)
            StrSQL = "SELECT TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName," & _
            "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
            StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
            StrSQL = StrSQL + " Where TblItems.ItemID=" & LngItemID
            Set Rs = New ADODB.Recordset
            Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (Rs.BOF Or Rs.EOF) Then
                .LblData(7).Caption = IIf(IsNull(Rs("ItemCode").Value), "", Rs("ItemCode").Value)
                .LblData(6).Caption = IIf(IsNull(Rs("ItemName").Value), "", Rs("ItemName").Value)
                .LblData(5).Caption = IIf(IsNull(Rs("GroupName").Value), "", Rs("GroupName").Value)
                .LblData(4).Caption = IIf(IsNull(Rs("SallingPrice").Value), "", Rs("SallingPrice").Value)
                .LblData(8).Caption = IIf(IsNull(Rs("CustomerPrice").Value), "", Rs("CustomerPrice").Value)
                .LblData(13).Caption = IIf(IsNull(Rs("DealerPrice").Value), "", Rs("DealerPrice").Value)
            End If
            If Me.GridTrans = InvoiceTransaction Then
                ItemTransInfo = GetLastItemTrans(LngItemID, 2)
                .lbl(4).Caption = "√Œ— ›« Ê—… »Ì⁄ ··’‰›:"
                .lbl(10).Caption = "«”„ «·⁄„Ì·:-"
            ElseIf Me.GridTrans = PurchaseTransaction Then
                ItemTransInfo = GetLastItemTrans(LngItemID, 1)
                .lbl(4).Caption = "√Œ— ›« Ê—… ‘—«¡ ··’‰›:"
                .lbl(10).Caption = "«”„ «·„Ê—œ:-"
            End If
            .LblData(10).Caption = ItemTransInfo.TransactionSerial
            .LblData(9).Caption = ItemTransInfo.TransactionDate
            .LblData(11).Caption = ItemTransInfo.SngItemPrice
            .LblData(12).Caption = ItemTransInfo.StrCustomerName
            
            
            .left = lLeft
            .top = LTop
            Rs.Close
            Set Rs = Nothing
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
        Set .ImageList = MDIFrmMain.ImgLstTree
        
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
Dim I As Long
'Dim LngNewEmptyRow As Long

For I = Me.Grid.FixedRows To Me.Grid.Rows - 1
    If Val(Me.Grid.TextMatrix(I, Grid.ColIndex("Name"))) = 0 Then
        NewEmptyRow = I
        Exit Function
    End If
Next I
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
Dim I As Integer
Dim Rs As ADODB.Recordset
Dim LngItemID As Long
Dim FG As VSFlex8UCtl.vsFlexGrid
Dim LngFindRow As Long

If Not m_FrmItemsNewPrices Is Nothing Then
    ShowNewItemsPrice = False
    Exit Function
End If

Set m_FrmItemsNewPrices = New FrmItemsNewPrices
Set FG = m_FrmItemsNewPrices.FG
FG.Rows = FG.FixedRows

Set Rs = New ADODB.Recordset
Rs.Open "TblItems", Cn, adOpenStatic, adLockReadOnly, adCmdTable
If Not Rs.BOF Or Rs.EOF Then
    Rs.MoveFirst
    For I = 1 To Me.Grid.Rows - 1
        LngItemID = Val(Me.Grid.TextMatrix(I, Me.Grid.ColIndex("Name")))
        If LngItemID <> 0 Then
            LngFindRow = FG.FindRow(LngItemID, FG.FixedRows, FG.ColIndex("ItemID"), False, True)
            If LngFindRow = -1 Then
                Rs.MoveFirst
                Rs.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
                If Not (Rs.BOF Or Rs.EOF) Then
                    With FG
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("Serial")) = .Rows - 1
                        .TextMatrix(.Rows - 1, .ColIndex("ItemID")) = IIf(IsNull(Rs("ItemID").Value), "", Rs("ItemID").Value)
                        .TextMatrix(.Rows - 1, .ColIndex("ItemCode")) = IIf(IsNull(Rs("ItemCode").Value), "", Rs("ItemCode").Value)
                        .TextMatrix(.Rows - 1, .ColIndex("ItemName")) = IIf(IsNull(Rs("ItemName").Value), "", Rs("ItemName").Value)
                        .TextMatrix(.Rows - 1, .ColIndex("SallingPrice")) = IIf(IsNull(Rs("SallingPrice").Value), "", Rs("SallingPrice").Value)
                        .TextMatrix(.Rows - 1, .ColIndex("CustomerPrice")) = IIf(IsNull(Rs("CustomerPrice").Value), "", Rs("CustomerPrice").Value)
                        .TextMatrix(.Rows - 1, .ColIndex("DealerPrice")) = IIf(IsNull(Rs("DealerPrice").Value), "", Rs("DealerPrice").Value)
                    End With
                End If
            End If
        End If
    Next I
End If
FG.AutoSize 0, FG.Cols - 1, False
m_FrmItemsNewPrices.Show
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
Dim StrItemSize As String

'-------------------------------------------------
'Start
'-------------------------------------------------
CheckDataEntered = False
'«·ﬂ‘› ⁄·Ï «‰ «·›« Ê—…  Õ ÊÏ ⁄·Ï √’‰«›
If Me.Grid.FixedRows = Me.Grid.Rows Or _
    ModFgLib.ItemsInGrid(Me.Grid, Me.Grid.ColIndex("Code")) = -1 Then
    Msg = "·« ÊÃœ √’‰«› ›Ï «·›« Ê—… ...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Screen.MousePointer = vbDefault
    Exit Function
End If
If Me.RestrictedAssbliedItemID <> 0 Then
    If CheckRestrictedItem = False Then
        CheckDataEntered = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
End If
With Me.Grid
    For LngRowNum = Me.Grid.FixedRows To Me.Grid.Rows - 1
        If Val(.TextMatrix(LngRowNum, .ColIndex("Name"))) <> 0 Then
            StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
            '----------«· «ﬂœ „‰ ≈œŒ«· ﬂ„Ì… «·’‰›
            If Val(.TextMatrix(LngRowNum, .ColIndex("Count"))) <= 0 Then
                Msg = "„‰ ›÷·ﬂ ÌÃ»  ÕœÌœ «·ﬂ„Ì… «·Œ«’… »«·’‰›" & Chr(13)
                Msg = Msg + StrCurrentItemName & Chr(13)
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                FocusGridOnCell LngRowNum, Grid.ColIndex("Count")
                Screen.MousePointer = vbDefault
                Exit Function
            '-------«· «ﬂœ „‰  ≈œŒ«· ”⁄— «·’‰›
            ElseIf Val(.TextMatrix(LngRowNum, .ColIndex("Price"))) <= 0 Then
                If Grid.ColIndex("DiscountType") <> -1 Then
                    If Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) <> 4 Then
                        Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
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
                    Msg = " ÌÃ»  ÕœÌœ À„‰ «·’‰›" & Chr(13)
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
                    Msg = "„‰ ›÷·ﬂ Õœœ ‰Ê⁄ «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                    Msg = Msg + StrCurrentItemName & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.Row = LngRowNum
                    Grid.Col = Grid.ColIndex("DiscountType")
                    Grid.ShowCell LngRowNum, Grid.ColIndex("DiscountType")
                    Grid.SetFocus
                    Exit Function
                '-------«· «ﬂœ „‰ ≈œŒ«· ﬁÌ„… «·Œ’„ „Ÿ»Êÿ
                ElseIf (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 2 Or _
                    (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = 3 Then
                    If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal"))) = "" Then
                        Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        FocusGridOnCell LngRowNum, Grid.ColIndex("DiscountVal")
                        Screen.MousePointer = vbDefault
                        Exit Function
                    ElseIf Not IsNumeric((Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountVal")))) Then
                        Msg = "ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & Chr(13)
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        Msg = Msg + "€Ì— —ﬁ„Ì…"
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
                        Msg = "«·’‰› " & Chr(13)
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        Msg = Msg + "·Â ”Ì—Ì«·" & Chr(13)
                        Msg = Msg + "„‰ ›÷·ﬂ √œŒ· «·”Ì—Ì«· «·Œ«’ »Â"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        FocusGridOnCell LngRowNum, Grid.ColIndex("Serial")
                        Screen.MousePointer = vbDefault
                        Exit Function
                    '›Ï ‰Ÿ«„ «·”Ì—Ì«· ·«»œ «‰ ÌﬂÊ‰ ﬂ· —ﬁ„ ›Ï Œ«‰… „‰›’·…
                    ElseIf Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Count"))) > 1 Then
                        Msg = "«·’‰› " & Chr(13)
                        Msg = Msg + StrCurrentItemName & Chr(13)
                        Msg = Msg + "·Â ”Ì—Ì«·" & Chr(13)
                        Msg = Msg + "·« Ì„ﬂ‰  ”ÃÌ· √ﬂÀ— „‰ ﬁÿ⁄… Ê«Õœ…" & Chr(13)
                        Msg = Msg + "Ì„ﬂ‰ﬂ  ”ÃÌ· ﬂ· ﬁÿ⁄… ›Ì ’› „” ﬁ·"
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
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + _
                        vbMsgBoxRtlReading, App.Title
                        Grid.Row = LngRowNum
                        Grid.Col = Grid.ColIndex("Serial")
                        Grid.ShowCell LngRowNum, Grid.ColIndex("Serial")
                        Grid.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Function
                    ElseIf .ColIndex("guaranteeTime") <> -1 Then
                        If .TextMatrix(LngRowNum, .ColIndex("guaranteeTime")) <> "" Then
                            If Not (IsNumeric(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) _
                            Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) < 1 _
                            Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) > 36 _
                            Or Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) <> _
                            Int(Val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Then
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
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + _
                vbMsgBoxRtlReading, App.Title
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
                    If Me.TxtModFlag.text = "N" Then
                        Set RsRepeat = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, , , , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                    ElseIf Me.TxtModFlag.text = "E" Then
                        Set RsRepeat = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, , _
                        Val(Me.TxtInvID.text), , , Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))))
                    End If
                    If Not (RsRepeat.EOF Or RsRepeat.BOF) Then
                        If RsRepeat("Qty").Value > 0 Then
                            Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                            Msg = Msg + StrCurrentItemName & Chr(13)
                            Msg = Msg + "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                            Msg = Msg + "„ÊÃÊœ ‰›” «·”Ì—Ì«· ·œÌ‰« ›Ï «·„Œ“‰ "
                            If Not StoreName Is Nothing Then
                                Msg = Msg & "[" & StoreName.text & "]"
                            End If
                            Msg = Msg & Chr(13) & "------------------------------------"
                            Msg = Msg & Chr(13) & "Õ«Ê· «‰  ” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ ”Ì—Ì«· ·„⁄—›… «·„“Ìœ"
                            Msg = Msg & Chr(13) & "Â·  —Ìœ «‰ ÌﬁÊ„ «·»—‰«„Ã »«·»ÕÀ ·ﬂ ⁄‰  Â–« «·”Ì—Ì«· «·¬‰...!!!"
                            IntRes = MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + _
                            vbMsgBoxRtlReading, App.Title)
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
    If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = Destruction) Then
        If SystemOptions.SysAllowStockNegative = False Then
            For LngRowNum = 1 To Grid.Rows - 1
                If Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType"))) = 0 Then '·«»œ Ê«‰ ÌﬂÊ‰ «·’‰› ”·⁄…
                If Grid.TextMatrix(LngRowNum, Grid.ColIndex("Name")) <> "" Then
                    '-------------------------------------------------------------------
                    '«·»Ì«‰«  «·√”«”Ì… —ﬁ„ «·’‰› Ê «”„ «·’‰› Ê «·”Ì—Ì«·
                    LngItemID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Name")))
                    StrCurrentItemName = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
                    StrCurrentItemCode = Grid.Cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Code"))
                    LngColorID = Val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ColorID")))
                    StrItemSize = Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemSize"))
                    
                    If Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) <> "" Then
                        StrItemSerial = Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")))
                    Else
                        StrItemSerial = ""
                    End If
                    DblOutQty = CheckTotalQty(LngRowNum)
                    '-------------------------------------------------------------------
                    If Me.TxtModFlag.text = "N" Then
                        Set RsTest = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, _
                        Me.DtpBillDate.Value, , , , StrItemSerial, True, LngColorID, StrItemSize)
                    ElseIf Me.TxtModFlag.text = "E" Then
                        '·«ÕŸ «‰Â ›Ï Õ«·… «· ⁄œÌ· ·«»œ „‰ ≈—”«· —ﬁ„ «·›« Ê—… ··œ«·…
                        Set RsTest = GetItemQuantityStock(LngItemID, Me.StoreName.BoundText, _
                        Me.DtpBillDate.Value, Val(Me.TxtInvID.text), , , StrItemSerial, True, LngColorID, StrItemSize)
                    End If
                    'Stop here
                    Screen.MousePointer = vbDefault
                    If RsTest.EOF Or RsTest.BOF Then
                        Msg = "«·’‰› : " & StrCurrentItemName & Chr(13)
                        Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                        Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                        Msg = Msg & "·« ÊÃœ „‰Â «Ï ﬂ„Ì… ›Ï «·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                        Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                        Msg = Msg & "--------------------------" & Chr(13) & ""
                        Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        FocusGridOnCell LngRowNum, .ColIndex("Name")
                        TerminateRecordset RsTest
                        Exit Function
                    ElseIf (RsTest("Qty").Value < DblOutQty) And StrItemSerial = "" Then
                        Msg = "«·’‰› : " & StrCurrentItemName & Chr(13)
                        Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                        Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                        Msg = Msg & "≈Ã„«·Ï «·ﬂ„Ì… «·„ÊÃÊœ… „‰Â ›Ï «·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                        Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                        Msg = Msg & "ÂÏ : " & RsTest("Qty").Value & Chr(13)
                        Msg = Msg & "·«Ì„ﬂ‰ «·”„«Õ »Œ—ÊÃ Â–Â «·ﬂ„Ì….  " & Chr(13)
                        Msg = Msg & "--------------------------" & Chr(13) & ""
                        Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        FocusGridOnCell LngRowNum, .ColIndex("Count")
                        TerminateRecordset RsTest
                        Exit Function
                    ElseIf StrItemSerial <> "" Then
                        RsTest.MoveFirst
                        RsTest.Find "ItemSerial='" & StrItemSerial & "'", , adSearchForward, adBookmarkFirst
                        If RsTest.BOF Or RsTest.EOF Then
                            Msg = "«·”Ì—Ì«· : " & StrItemSerial & Chr(13)
                            Msg = Msg & "«·’‰› : " & StrCurrentItemName & Chr(13)
                            Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                            Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                            Msg = Msg & "Â–« «·”Ì—Ì«· €Ì— „ÊÃÊœ »«·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                            Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                            Msg = Msg & "--------------------------" & Chr(13) & ""
                            Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, .ColIndex("Count")
                            TerminateRecordset RsTest
                            Exit Function
                        ElseIf RsTest("Qty").Value <= 0 Then
                            Msg = "«·”Ì—Ì«· : " & StrItemSerial & Chr(13)
                            Msg = Msg & "«·’‰› : " & StrCurrentItemName & Chr(13)
                            Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & Chr(13)
                            Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & Chr(13)
                            Msg = Msg & "Â–« «·”Ì—Ì«· €Ì— „ÊÃÊœ »«·„Œ“‰ : " & Me.StoreName.text & Chr(13)
                            Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.Value) & Chr(13)
                            Msg = Msg & "--------------------------" & Chr(13) & ""
                            Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & Chr(13) & ""
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, .ColIndex("Count")
                            TerminateRecordset RsTest
                            Exit Function
                        End If
                    End If
                End If
                End If
            Next LngRowNum
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
Private Function CheckTotalQty(LngRow As Long) As Double
Dim I As Integer
Dim LngFoundGridRow As Long
Dim Dbl_OutQtyBySmallUnit As Double
If LngRow = 0 Then Exit Function
With Me.Grid
    LngFoundGridRow = 0
    For I = 1 To .Rows - 1
        LngFoundGridRow = .FindRow(.TextMatrix(LngRow, .ColIndex("Code")), I, .ColIndex("Code"), True, True)
        If LngFoundGridRow > 0 Then
            I = LngFoundGridRow
            Dbl_OutQtyBySmallUnit = Dbl_OutQtyBySmallUnit + Val(.TextMatrix(LngFoundGridRow, .ColIndex("Count")))
        ElseIf LngFoundGridRow = -1 Then        'did not found the item entered before
            Exit For
        End If
    Next I
End With
CheckTotalQty = Dbl_OutQtyBySmallUnit
End Function


Private Sub FocusGridOnCell(LngRow As Long, LngCol As Long)
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


Private Sub ItemComment(Row As Long, Col As Long)
Dim Frm As FrmGridAddItemComment
Set Frm = New FrmGridAddItemComment
With Me.Grid
    Frm.lbl(3).Caption = (Row + 1)
    Frm.lbl(4).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Code"))
    Frm.lbl(5).Caption = Me.Grid.Cell(flexcpTextDisplay, Row, .ColIndex("Name"))
    Set Frm.FG = Me.Grid
    Frm.LngRow = Row
    Frm.LngCol = Col
    Frm.TxtComment.text = Trim$(Me.Grid.TextMatrix(Row, Col))
    If Me.TxtModFlag.text = "R" Then
        Frm.TxtComment.Locked = True
        Frm.CmdOk.Enabled = False
    End If
End With
Frm.Show vbModal
End Sub

Private Sub PopUpSerialList(LngItemID As Long, Optional LngRow As Long = 0, _
Optional LngCol As Long = 0)

Dim Frm As FrmSerialList
Dim LngStoreID As Long
LngStoreID = Val(Me.StoreName.BoundText)
If LngItemID = 0 Or LngStoreID = 0 Then
    Exit Sub
End If
Set Frm = New FrmSerialList
Frm.RetrunType = 0
Frm.RetrunGridOnOneline = False
Set Frm.m_TextBox = Me.TxtSerial
Frm.GetData LngItemID, LngStoreID
'------------------------------------------
Frm.FgCol = LngCol
Frm.FgRow = LngRow
'------------------------------------------
Set Frm.xGrid = Me
Frm.Show vbModal

End Sub
Private Sub ClearTopDataEntry()
DCboItemCode.BoundText = ""
DcboItemName.BoundText = ""
CboItemCase.ListIndex = 0
If Me.GridTrans <> Template And Me.GridTrans <> ShowPrice Then
    TxtSerial.text = ""
End If
TxtQuantity.text = ""
TxtPrice.text = ""
CboItemCase.ListIndex = 0
End Sub
Private Sub TxtServiceTax_Change()
If Me.TxtModFlag.text = "E" Or Me.TxtModFlag.text = "N" Then
    Me.Calculate 1, , , True
End If
End Sub

Private Sub TxtServiceTax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtServiceTax.text, 0)
End Sub


Private Sub TxtStampTax_Change()
If Me.TxtModFlag.text = "E" Or Me.TxtModFlag.text = "N" Then
    Me.Calculate 1, , , True
End If
End Sub

Private Sub TxtStampTax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtStampTax.text, 0)
End Sub


Private Sub TxtTaxValue_Change()
If Me.TxtModFlag.text = "E" Or Me.TxtModFlag.text = "N" Then
    Me.Calculate 1, , , True
End If
End Sub

Private Sub TxtTaxValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTaxValue.text, 0)
End Sub
Public Function GetItemsTotal(m_ItemType As ItemsType) As Double
Dim DblTemp As Double
Dim RowNum As Long
Dim Msg  As String
On Local Error GoTo ErrTrap
With Me.Grid
    If m_ItemType = ItemsAllType Then
        For RowNum = .FixedRows To .Rows - 1
            DblTemp = DblTemp + (Val(.TextMatrix(RowNum, .ColIndex("Count"))) * Val(.TextMatrix(RowNum, .ColIndex("Price"))))
        Next RowNum
    ElseIf m_ItemType = ItemsGoodType Then
        For RowNum = .FixedRows To .Rows - 1
            If Val(.TextMatrix(RowNum, .ColIndex("ItemType"))) = 0 Then
                DblTemp = DblTemp + (Val(.TextMatrix(RowNum, .ColIndex("Count"))) * Val(.TextMatrix(RowNum, .ColIndex("Price"))))
            End If
        Next RowNum
    ElseIf m_ItemType = ItemsServiceType Then
        For RowNum = .FixedRows To .Rows - 1
            If .TextMatrix(RowNum, .ColIndex("ItemType")) = 1 Then
                DblTemp = DblTemp + (Val(.TextMatrix(RowNum, .ColIndex("Count"))) * Val(.TextMatrix(RowNum, .ColIndex("Price"))))
            End If
        Next RowNum
    End If
End With
GetItemsTotal = DblTemp
Exit Function
ErrTrap:
Msg = "ÕœÀ Œÿ« "
GetItemsTotal = 0
End Function

Public Function GetItemsCostTotal() As Double
Dim DblTemp As Double
Dim RowNum As Long
Dim Msg  As String
Dim DblRowCostTotal As Double

On Local Error GoTo ErrTrap
With Me.Grid
    For RowNum = .FixedRows To .Rows - 1
        DblRowCostTotal = Val(.TextMatrix(RowNum, .ColIndex("ItemCostPrice"))) * Val(.TextMatrix(RowNum, .ColIndex("Count")))
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

Private Function CheckRestrictedItemQty(LngWorkOrderID As Long, LngItemID As Long) As Boolean

Dim Msg As String
Dim I As Long
Dim RsOutQty As ADODB.Recordset
Dim Rs As ADODB.Recordset
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

StrSQL = "SELECT SUM(Qty) AS SumQty, ItemID, ItemCode, ItemName, StoreID, StoreName," & _
"GroupID, ColorID, ColorName, ItemSize "
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
StrSQL = "Select  * From TblUsers Where UserID=" & User_ID
Set RsUserOptions = New ADODB.Recordset
RsUserOptions.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (RsUserOptions.BOF Or RsUserOptions.EOF) Then
    If IsNull(RsUserOptions("AllowOverMax").Value) = True Then
        IntAllowMaxOverOut = -1
    ElseIf RsUserOptions("AllowOverMax").Value = 0 Then
        IntAllowMaxOverOut = 0
    ElseIf RsUserOptions("AllowOverMax").Value = 1 Then
        IntAllowMaxOverOut = 1
    End If
End If
'-----------------------------------------------------------------------------
StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & _
"dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty,dbo.TblItemsParts.TableId "
StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
StrSQL = StrSQL + " dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID "
StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID & ""
Set Rs = New ADODB.Recordset
Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

With Me.Grid
    For I = 1 To .Rows - 1
        SngOutQty = 0
        SngMaxOutQty = 0
        SngWillGoOutQty = 0
        
        LngSearchItemID = Val(.TextMatrix(I, .ColIndex("Name")))
        SngWillGoOutQty = Val(.TextMatrix(I, .ColIndex("Count")))
        If BolThereIsOutBefore = True Then
            RsOutQty.MoveFirst
            RsOutQty.Find "ItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst
            If Not (RsOutQty.BOF Or RsOutQty.EOF) Then
                SngOutQty = IIf(IsNull(RsOutQty("SumQty").Value), 0, RsOutQty("SumQty").Value)
            Else
                SngOutQty = 0
            End If
        Else
            SngOutQty = 0
        End If
        Rs.MoveFirst
        Rs.Find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst
        If Not (Rs.BOF Or Rs.EOF) Then
            SngMaxOutQty = Rs("PartItemQty").Value
            If (SngWillGoOutQty + SngOutQty) > SngMaxOutQty Then
                Msg = "«·’‰›..." & Rs("ItemName").Value
                Msg = Msg & Chr(13) & "√ﬁ’Ì ﬂ„Ì… „‰Â ›Ï «·’‰› «·„Ã„⁄ :" & SngMaxOutQty
                If IntAllowMaxOverOut = -1 Then
                    Msg = Msg & Chr(13) & "„·ÕÊŸ…: ·„ Ìﬁ„ „œÌ— «·‰Ÿ«„ »≈⁄œ«œ ’·«ÕÌ« ﬂ ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                    Msg = Msg & Chr(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CheckRestrictedItemQty = False
                    Exit Function
                ElseIf IntAllowMaxOverOut = 0 Then
                    Msg = Msg & Chr(13) & "„·ÕÊŸ…: ·Ì” ·ﬂ ’·«ÕÌ… ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                    Msg = Msg & Chr(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CheckRestrictedItemQty = False
                    Exit Function
                ElseIf IntAllowMaxOverOut = 1 Then
                    Msg = Msg & Chr(13) & "Â·  —Ìœ ’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄"
                    IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
                    If IntRes = vbNo Then
                        CheckRestrictedItemQty = False
                        Exit Function
                    End If
                End If
            End If
        End If
    Next I
End With
CheckRestrictedItemQty = True
End Function

Private Function CheckRestrictedItem() As Boolean

Dim StrSQL As String
Dim Rs As ADODB.Recordset
Dim Msg As String
Dim LngSearchItemID As Long
Dim StrTemp As String
Dim I As Long
Dim IntRes As Integer

StrSQL = "SELECT dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & _
"dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty,dbo.TblItemsParts.TableId "
StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
StrSQL = StrSQL + " dbo.TblItems ON dbo.TblItemsParts.PartItemID = dbo.TblItems.ItemID "
StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & Me.RestrictedAssbliedItemID & ""
Set Rs = New ADODB.Recordset
Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs.BOF Or Rs.EOF Then
    Msg = "«·’‰› «·„Õœœ ’‰› €Ì— „Ã„⁄ »—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    CheckRestrictedItem = False
    Exit Function
End If
With Me.Grid
    For I = .FixedRows To .Rows - 1
        LngSearchItemID = Val(.TextMatrix(I, .ColIndex("Name")))
        If LngSearchItemID <> 0 Then
            Rs.MoveFirst
            Rs.Find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst
            If Rs.BOF Or Rs.EOF Then
                StrTemp = "«·’‰› ..." & .Cell(flexcpTextDisplay, I, .ColIndex("Name"))
                Msg = StrTemp
                Msg = Msg & Chr(13) & "€Ì— „ÊÃÊœ ›Ï «·’‰› «·„Ã„⁄ «·„— »ÿ »Â «„— «·‘€·"
                Msg = Msg & Chr(13) & ""
                Msg = Msg & Chr(13) & "Â·  —Ìœ «·„ «»⁄… ›Ï Õ›Ÿ «„— «·‘€·...øøø"
                IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
                If IntRes = vbNo Then
                    CheckRestrictedItem = False
                    Exit Function
                End If
            End If
        End If
    Next I
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
    LngItemID = Val(.TextMatrix(LngRow, .ColIndex("Name")))
    LngUnitID = Val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
    
    If LngItemID = 0 Then Exit Sub
   
    StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & _
    "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & _
    "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & _
    "TblItemsUnits.FactorBySmallUnit "
    StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & _
    "TblUnites.UnitID"
    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
    StrSQL = StrSQL + " AND TblItemsUnits.UnitID=" & LngUnitID
     Set RsUnitData = New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    DblOrgQty = Val(.TextMatrix(LngRow, .ColIndex("Count")))
    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
        .TextMatrix(LngRow, .ColIndex("DefCount")) = DblOrgQty * RsUnitData("FactorByDefaultUnit").Value
        .TextMatrix(LngRow, .ColIndex("DefCount")) = Format(.TextMatrix(LngRow, .ColIndex("DefCount")), "0.000")
    End If
End With
RsUnitData.Close
Set RsUnitData = Nothing
End Sub

