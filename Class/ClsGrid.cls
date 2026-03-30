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
Dim Msg As String
Public WithEvents GridQuantity As VSFlexGrid
Attribute GridQuantity.VB_VarHelpID = -1
Public WithEvents Grid As VSFlexGrid
Attribute Grid.VB_VarHelpID = -1
Public WithEvents VatGrid As VSFlexGrid
Attribute VatGrid.VB_VarHelpID = -1
Public IsOrderOut As Boolean
Public IsNotSave As Boolean
Public WithEvents GrdItems As VSFlexGrid
Attribute GrdItems.VB_VarHelpID = -1
Public WithEvents GrdItems2 As VSFlexGrid
Attribute GrdItems2.VB_VarHelpID = -1
Public WithEvents LblValueAdded As Label
Attribute LblValueAdded.VB_VarHelpID = -1
Public mTaxTobacco As Double
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
Dim StrComboList As String
Public customerid As Long
Public TxtValueAdded As TextBox
Public WithEvents Customer As DataCombo
Attribute Customer.VB_VarHelpID = -1
Dim IsDuplicateItem As Boolean
Public WithEvents cmbGroup As DataCombo
Attribute cmbGroup.VB_VarHelpID = -1
Public WithEvents cmbGroup1 As DataCombo
Attribute cmbGroup1.VB_VarHelpID = -1
Public WithEvents cmbGroup2 As DataCombo
Attribute cmbGroup2.VB_VarHelpID = -1
Dim rsDummy As ADODB.Recordset

Public WithEvents Branch As DataCombo
Attribute Branch.VB_VarHelpID = -1
Dim cCompanyInfo As New ClsCompanyInfo
Dim rsAccount As ADODB.Recordset
Public WithEvents CboDiscount_Type As ComboBox
Attribute CboDiscount_Type.VB_VarHelpID = -1
Public SelectedLotNo As String
Public WithEvents TxtDiscount_Val As TextBox
Attribute TxtDiscount_Val.VB_VarHelpID = -1

Public txtPumbDet1 As TextBox
Attribute txtPumbDet1.VB_VarHelpID = -1
Public txtPumbDet2 As TextBox

Public lblTotalSalesPrice As Label
Public txt_Currency_rate As TextBox
Dim mMainRow As Long
Public WithEvents cmdInsert As CommandButton
Attribute cmdInsert.VB_VarHelpID = -1
Public TxtValueCash As TextBox
Public WithEvents XPTxtDiscountVal As TextBox
Attribute XPTxtDiscountVal.VB_VarHelpID = -1

Dim mValue As Double
Public TxtValueDelay As TextBox
Dim s As String
Public XPDtbBill As DTPicker

Public TxtValuechque As TextBox
Public DefStorePurchase As Integer
Public mStorePurName As String
Public WithEvents TxtTaxValue As TextBox
Attribute TxtTaxValue.VB_VarHelpID = -1

Public WithEvents TxtAddTax As TextBox
Attribute TxtAddTax.VB_VarHelpID = -1

Public WithEvents txtCallingID As TextBox
Attribute txtCallingID.VB_VarHelpID = -1

Public WithEvents TxtStampTax As TextBox
Attribute TxtStampTax.VB_VarHelpID = -1

Public WithEvents TxtServiceTax As TextBox
Attribute TxtServiceTax.VB_VarHelpID = -1
Public TxtFillData As TextBox
Public FramePay As frame
'NewWael
Public frmname As String

Public WithEvents StoreName As DataCombo
Attribute StoreName.VB_VarHelpID = -1

Public LblInvProfit As Label

Public LblItemsCount As Label

Public WithEvents DtpBillDate As DTPicker
Attribute DtpBillDate.VB_VarHelpID = -1

Public TxtInvID As TextBox

Public TxtNots As TextBox
 
Public WithEvents txtManulaVat  As VB.TextBox
Attribute txtManulaVat.VB_VarHelpID = -1
Private FrmComment As FrmFlexQtyTip

Public WithEvents m_FrmItemsNewPrices As Form
Attribute m_FrmItemsNewPrices.VB_VarHelpID = -1

Public WithEvents LblTotalAll As Label
Attribute LblTotalAll.VB_VarHelpID = -1
Public WithEvents TxtItemCodeB    As TextBox
Attribute TxtItemCodeB.VB_VarHelpID = -1
Public WithEvents TxtItemCodeB1    As TextBox
Attribute TxtItemCodeB1.VB_VarHelpID = -1
Public WithEvents txtItemCodeSearch As TextBox
Attribute txtItemCodeSearch.VB_VarHelpID = -1
Public WithEvents txtItemCodeSearch2 As TextBox
Attribute txtItemCodeSearch2.VB_VarHelpID = -1
Public WithEvents LblDiscountsTotal As Label
Attribute LblDiscountsTotal.VB_VarHelpID = -1

Public WithEvents LblCommision As Label
Attribute LblCommision.VB_VarHelpID = -1

Public WithEvents LblTotalQty As Label
Attribute LblTotalQty.VB_VarHelpID = -1

Public LblTaxSalesValue As Label

Public LblTaxAddValue As Label

Public LblTaxStampValue As Label
Public mPumpId As Long

Public LblTaxServiceValue As Label
Public LBLGross As Label

'«·Ã“¡ «·Œ«’ » ⁄»∆… »Ì«‰«  «·√’‰«›
Public WithEvents DCboItemCode As DataCombo  'ﬂÊœ «·’‰›
Attribute DCboItemCode.VB_VarHelpID = -1

Public WithEvents DCboItemName As DataCombo  '«”„ «·’‰›
Attribute DCboItemName.VB_VarHelpID = -1
Public WithEvents cmbPumbType As DataCombo  '«”„ «·’‰›
Attribute cmbPumbType.VB_VarHelpID = -1
Public WithEvents cmbAccounts As DataCombo  '«”„ «·’‰›
Attribute cmbAccounts.VB_VarHelpID = -1

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


Public WithEvents chkIsTransfere As CheckBox
Attribute chkIsTransfere.VB_VarHelpID = -1
Public WithEvents chkTaxExempt As CheckBox
Attribute chkTaxExempt.VB_VarHelpID = -1


Public RSPH As DataCombo    '«”„ «·’‰›
Public RCLY As DataCombo    '«”„ «·’‰›
Public RAxis As DataCombo    '«”„ «·’‰›
Public LSPH As DataCombo    '«”„ «·’‰›
Public LCLY As DataCombo    '«”„ «·’‰›
Public LAxis As DataCombo    '«”„ «·’‰›

Public txtLCL As TextBox
Public txtRCL As TextBox



Public RIPD As DataCombo
Public LIPD As DataCombo
Public RADD As DataCombo
Public LADD   As DataCombo
Public RSH As DataCombo
Public LSH As DataCombo
Public RPRISM As DataCombo
Public LPRISM As DataCombo

Dim sComboStore As String
'Public TxtGuarante As TextBox  '„œ… «·÷„«‰
'Public CboDiscountType As  ComboBox    '‰Ê⁄ «·Œ’„
'Public TxtDiscountVal As TextBox  'ﬁÌ„… «·Œ’„
Private m_GridTrans As GridTransType
Dim cSearchDcbo(3) As clsDCboSearch
Dim ACCStrComboList As String
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
                            Dim FromDate As Date
                        Dim AverageMonthly As Double
                        Dim ItemBalance As Double
                         Dim FirstPeriodDateInthisYear  As Date

Private QtyFromBarCode As String
Private CodeFromBarCode As String



Public Property Get GridTrans2() As GridTransType
    GridTrans2 = m_GridTrans2
End Property




Private Sub Customer_GotFocus()
  'Add
    Customer.backcolor = vbYellow

End Sub

Private Sub Customer_KeyDown(KeyCode As Integer, Shift As Integer)
 
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub Customer_LostFocus()
   
    Customer.backcolor = vbWhite
End Sub


Private Sub TxtCustCode_GotFocus()


TxtCustCode.backcolor = vbYellow
End Sub

Private Sub TxtCustCode_KeyDown(KeyCode As Integer, Shift As Integer)
 
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub TxtCustCode_LostFocus()
 
 TxtCustCode.backcolor = vbWhite
End Sub

  
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
If m_GridTrans2 = 2 And UCase(frmname) = "FRMSALEBILL6" Then
    CheckMaultyItem = False
    Exit Function
End If
CheckMaultyItem = False
With Grid
For i = 1 To .rows - 1
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
Function GetLocationID(Optional ItemID As Double) As Double
If ItemID = 0 Then Exit Function
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT        MAX(locid) AS LocaID"
sql = sql & " From dbo.TblItemLoc"
sql = sql & " Where (ItemID = " & ItemID & ")"
sql = sql & " GROUP BY ItemID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetLocationID = IIf(IsNull(rs2("LocaID").value), 0, rs2("LocaID").value)
Else
GetLocationID = 0
End If
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
If Number < Grid.rows Then
   If Grid.rows > 2 And Grid.TextMatrix(Number, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Number
    
    End If
 End If
Loop While Number < Grid.rows
   If Grid.rows - 1 > 1 And Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("Code")) = "" Then
    Grid.RemoveItem Grid.rows - 1
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
   
        Sendkeys "{TAB}"
    End If
End Sub
 
 
 Private Sub CboItemCase_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then
   
        Sendkeys "{TAB}"
    End If

End Sub
Sub ColorExperDate()
RemoveItems
 Dim i As Integer
If SystemOptions.AllowExperDateFIFO = True Then
With Grid
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("Name")) <> "" Then
If .TextMatrix(i, .ColIndex("ExpiryDate")) = "" Then
.cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
ElseIf DateDiff("d", .TextMatrix(i, .ColIndex("ExpiryDate")), Date) >= 0 Then
.cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
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
Private Sub cmdInsert_Click()

Dim s As String
Dim i As Long


   For i = 1 To GrdItems.rows - 1
   
        If GrdItems.ValueMatrix(i, GrdItems.ColIndex("IsSelected")) Then
            ' DCboItemName.BoundText = GrdItems.TextMatrix(GrdItems.Row, GrdItems.ColIndex("ItemID"))
            If Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("Code")) <> "" Then
                Grid.rows = Grid.rows + 1
            End If
            
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
                TxtQuantity.Text = 1
            End If
            If GrdItems.TextMatrix(i, GrdItems.ColIndex("TypeFrame")) = "R" Then
                RSPH = GrdItems.TextMatrix(i, GrdItems.ColIndex("SPHT"))
                RCLY = GrdItems.TextMatrix(i, GrdItems.ColIndex("CLYT"))
                LCLY = ""
                LSPH = ""
                
                
            
            End If
            If GrdItems.TextMatrix(i, GrdItems.ColIndex("TypeFrame")) = "L" Then
                LSPH = GrdItems.TextMatrix(i, GrdItems.ColIndex("SPHT"))
                LCLY = GrdItems.TextMatrix(i, GrdItems.ColIndex("CLYT"))
                RSPH = ""
                RCLY = ""
            End If
            mMainRow = val(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemID2")))
            AddItemInGrid Grid.rows - 1, val(GrdItems.TextMatrix(i, GrdItems.ColIndex("ItemID"))), 1, val(Me.TxtPrice.Text), , Trim$(Me.TxtSerial.Text), , , , , , , , , , , , , , 1
            'CmdAddData_Click
            
        End If
            
    Next
   i = 1
   For i = 1 To GrdItems2.rows - 1
   
        If GrdItems2.ValueMatrix(i, GrdItems2.ColIndex("IsSelected")) Then

        
       ' DCboItemName.BoundText = GrdItems.TextMatrix(GrdItems.Row, GrdItems.ColIndex("ItemID"))
        If Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("Code")) <> "" Then
            Grid.rows = Grid.rows + 1
        End If
                    
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
                TxtQuantity.Text = 1
            End If

      
        If GrdItems2.TextMatrix(i, GrdItems2.ColIndex("TypeFrame")) = "R" Then
            RSPH = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("SPHT"))
            RCLY = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("CLYT"))
            RADD = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ADD"))
            
            RCLY = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("CLYT"))
            
            RADD.Text = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ADD"))
            

            Grid.TextMatrix(mMainRow, Grid.ColIndex("ReadSPH")) = val(RSPH)
            Grid.TextMatrix(mMainRow, Grid.ColIndex("ReadCYL")) = RCLY.Text

                
                
                                

            LADD = ""
            LADD.Text = ""
            LCLY = ""
            LSPH = ""

        End If
        If GrdItems2.TextMatrix(i, GrdItems2.ColIndex("TypeFrame")) = "L" Then
            LSPH = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("SPHT"))
            LCLY = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("CLYT"))
            LADD = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ADD"))
            LADD.Text = GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ADD"))
            
            Grid.TextMatrix(mMainRow, Grid.ColIndex("LReadSPH")) = LSPH
            Grid.TextMatrix(mMainRow, Grid.ColIndex("LReadCYL")) = LCLY.Text


            RADD.Text = ""
            RADD = ""
            
            RSPH = ""
            RCLY = ""
        End If
        mMainRow = val(GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ItemID2")))
        AddItemInGrid Grid.rows - 1, val(GrdItems2.TextMatrix(i, GrdItems2.ColIndex("ItemID"))), 1, val(Me.TxtPrice.Text), , Trim$(Me.TxtSerial.Text), , , , , , , , , , , , , , 1
        'CmdAddData_Click
        
    End If
   Next
   CheckQty
 
    FramePay.Visible = False


End Sub

 
Private Sub CmdAddData_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyReturn Then
        
        If txtItemCodeSearch Is Nothing Then
            
            DCboItemCode.SetFocus
        Else
            txtItemCodeSearch.SetFocus
        End If
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
    Dim LngUnitID As Long
    Dim StrItemDataMsg As String
    Dim mIsGroup As Boolean
    If cmbGroup Is Nothing Then
        mIsGroup = False
    Else
        If cmbGroup.Text <> "" Then
            mIsGroup = True
        End If
    End If
    
    If UCase(frmname) <> "FRMSALEBILL3" Then
    
    If Not mIsGroup Then
        If DCboItemCode.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
            Else
                Msg = "Enter item Code"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboItemCode.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        
        
        If DCboItemName.Text = "" Then
  '      Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        
                                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
                        Else
                        Msg = "Enter item name"
                        End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemName.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If
    Else
    GoTo mLenses
    End If
    End If
     

    If Me.GridTrans2 = InvoiceTransactionCompose Then
    If Customer.Text = "" Then
  '      Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        
                                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ «”„ «·⁄„Ì·"
                        Else
                        Msg = "Enter customer name"
                        End If


        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Customer.SetFocus
        Sendkeys "{F4}"
        Exit Sub
        End If
    End If
    If UCase(frmname) = "FRMSALEBILL3" Or UCase(frmname) = "FRMSALEBILL2" Then
        CboItemCase.ListIndex = 1
    End If
    
If CboItemCase.ListIndex = -1 Then
CboItemCase.ListIndex = 1
End If
    
    If CboItemCase.ListIndex = -1 Then
        
       ' Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ Õ«·… «·’‰›"
                        Else
                        Msg = "Enter item Status"
                        End If


        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboItemCase.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If

    If Not TxtSerial Is Nothing Then
        If Trim(Me.TxtSerial.Text) <> "" Then

            With Me.Grid
                StrSerial = Trim(Me.TxtSerial.Text)
                LngRow = .FindRow(StrSerial, .FixedRows, .ColIndex("Serial"), False, True)

                If LngRow <> -1 Then
                    Msg = "Â–« «·”Ì—Ì«· „œŒ· „”»ﬁ« ›Ï «·”ÿ— —ﬁ„ " & LngRow
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If

            End With

        End If
    End If
    

If Not mIsGroup Then
If (val(TxtQuantity.Text) = 0 And Not RSPH Is Nothing And (DCboItemCode.Text <> "" Or DCboItemName.Text <> "")) Then
        
        
                    
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
                TxtQuantity.Text = 1
            End If

End If
    If val(TxtQuantity.Text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰›"
                        Else
                        Msg = "Enter item Qty"
                        End If
                        
   '     Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtQuantity.SelStart = 0
            TxtQuantity.SelLength = Len(TxtQuantity)
            TxtQuantity.SetFocus

        Exit Sub
    End If
Else
            
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
                TxtQuantity.Text = 1
            End If

End If
    If Not IsNumeric(TxtQuantity.Text) Then
      '  Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        
                    If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                        Else
                        Msg = "Qty Must be Numeric"
                        End If
                        
                        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtQuantity.SelStart = 0
            TxtQuantity.SelLength = Len(TxtQuantity)
            TxtQuantity.SetFocus

        Exit Sub
    End If
    
    
    
    If Not mIsGroup Then
        If TxtPrice.Text = "" Then
            'Msg = "ÌÃ»  ÕœÌœ «·”⁄—"
            
                         If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "ÌÃ»  ÕœÌœ «·”⁄— "
                            Else
                            Msg = "Price Must be Entered"
                            End If
                            
                            
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtPrice.SetFocus
            Exit Sub
        End If
    Else
    TxtPrice = 0
    End If
    If Not IsNumeric(TxtPrice.Text) Then
        Msg = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                     If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "«·”⁄— ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                        Else
                        Msg = "Price Must be Numeric"
                        End If
                        
                        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
                        
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    
    If UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
       ' s = "Select tblItems.* from tblItems where itemid In (Select tblPumpType.ItemId from tblPumpType)"
       ' Set RsItems = New ADODB.Recordset
       ' RsItems.Open s, Cn, adOpenStatic, adLockReadOnly
    End If
    
    RsItems.Find "ItemID=" & LngItemID, , adSearchForward, adBookmarkFirst

    
If Not mIsGroup Then
        If RsItems.BOF Or RsItems.EOF Then
           ' Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ"
            
                                         If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ"
                            Else
                            Msg = "Invalid Item Code"
                            End If
                            
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        
    ElseIf (RsItems("AssbliedItem").value = True Or RsItems("AssbliedItem").value = -1 Or RsItems("AssbliedItem").value = 1) Then
        '’‰› „Ã„⁄
    
    ElseIf (RsItems("ItemType").value = 1) Then
        '’‰› Œœ„…
    
    Else
        StrItemDataMsg = "«·’‰› : " & RsItems("ItemName").value
        StrItemDataMsg = StrItemDataMsg & CHR(13) & "—ﬁ„:" & RsItems("ItemID").value
        StrItemDataMsg = StrItemDataMsg & CHR(13) & "ﬂÊœ:" & RsItems("ItemCode").value
        
        GetDefaultItemUnit LngItemID, LngUnitID
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
End If
    With Grid

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If

        If Not Me.TxtSerial Is Nothing Then
            If Not DtpBillDate Then
                If Not RSPH Is Nothing Then
                     
                    mMainRow = .rows - 1
                    AddItemInGrid .rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.Text), , LngUnitID, , , , , , , , , DtpBillDate.value, , , 1
                Else
                    AddItemInGrid .rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.Text), , LngUnitID, , , , , , , , , DtpBillDate.value
                End If
                
                
            Else
                If Not RSPH Is Nothing Then
                    mMainRow = .rows - 1
                    AddItemInGrid .rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.Text), , , , , , , , , , , , , , 1
                Else
                    AddItemInGrid .rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, Trim$(Me.TxtSerial.Text)
                End If
                    
            End If
            ' customer_screen.bill_no.Caption = frmsalebill.XPTxtBillID.text ' TxtTransSerial.text 'salim2
            'customer_screen.List1.AddItem DcboItemName.BoundText & Chr(9) & DcboItemName.text & "" & Val(TxtQuantity.text) & "*" & Me.TxtPrice.text
            'customer_screen.List2.AddItem Val(TxtQuantity.text) * Me.TxtPrice.text
            '  customer_screen.total
            
mLenses:
    If Not RSPH Is Nothing Then
            If RSPH <> "" Then
                
            If FramePay.Visible = False Then
                Dim s As String
                Dim rsDummy As New ADODB.Recordset
                s = " SELECT  ItemID2 = " & mMainRow & ",  TypeFrame ='R', ti.ItemCode,ti.ItemID,ti.ItemName,ti.SphereID,ti.CylinderID,g.GroupName,tu.UnitName,SPHTable.SPHT,CLYTable.CLYT"
                s = s & "      FROM TblItems AS ti INNER JOIN Groups AS g ON g.GroupID = ti.GroupID"
                s = s & " INNER JOIN TblItemsUnits ON TblItemsUnits.ItemID = ti.ItemID"
                s = s & " INNER JOIN TblUnites AS tu ON tu.UnitID = TblItemsUnits.UnitID"
                s = s & " INNER JOIN SPHTable ON ti.SphereID = SPHTable.ID"
                s = s & " INNER JOIN CLYTable ON ti.CylinderID = CLYTable.ID"
                s = s & " Where  (SPHTable.ID = " & val(GetSphCylID(0, RSPH))
                s = s & " and CLYTable.ID= " & val(GetSphCylID(1, RCLY)) & ")"
                     
                If (DCboItemCode.Text = "" And DCboItemName.Text = "") Then
                   s = s & " and ti.LensesTypesID  = " & val(cmbGroup1.BoundText)
                   s = s & " and ti.GroupID = " & val(cmbGroup.BoundText)
                End If
                
                
                If LSPH <> "" Then
                    s = s & " Union All                                "
                    
                    s = s & " SELECT ItemID2 = " & mMainRow & ",TypeFrame ='L', ti.ItemCode,ti.ItemID,ti.ItemName,ti.SphereID,ti.CylinderID,g.GroupName,tu.UnitName,SPHTable.SPHT,CLYTable.CLYT"
                    s = s & "      FROM TblItems AS ti INNER JOIN Groups AS g ON g.GroupID = ti.GroupID"
                    s = s & " INNER JOIN TblItemsUnits ON TblItemsUnits.ItemID = ti.ItemID"
                    s = s & " INNER JOIN TblUnites AS tu ON tu.UnitID = TblItemsUnits.UnitID"
                    s = s & " INNER JOIN SPHTable ON ti.SphereID = SPHTable.ID"
                    s = s & " INNER JOIN CLYTable ON ti.CylinderID = CLYTable.ID"
                    s = s & " Where  (SPHTable.ID = " & val(GetSphCylID(0, LSPH))
                    s = s & " and CLYTable.ID= " & val(GetSphCylID(1, LCLY)) & ")"
                     
                     If (DCboItemCode.Text = "" And DCboItemName.Text = "") Then
                        s = s & " and ti.LensesTypesID  = " & val(cmbGroup1.BoundText)
                        s = s & " and ti.GroupID = " & val(cmbGroup.BoundText)
                     End If
                    
              End If
                

                loadgrid s, GrdItems, True, False
                If GrdItems.rows > 1 Then
                    FramePay.Visible = True
                    'Exit Sub
                    GrdItems.left = cmdInsert.Width
                    GrdItems2.left = cmdInsert.Width
                    GrdItems.Height = FramePay.Height - 200
                    GrdItems.Width = FramePay.Width - 150 - cmdInsert.Width

                End If
            End If
            
        End If
        
        If Not RADD Is Nothing And (DCboItemCode.Text <> "" Or DCboItemName.Text <> "") And SystemOptions.IsShowLensesDetails Then
            If RADD <> "" Or LADD <> "" Then
                   s = " SELECT ItemID2 = " & mMainRow & ",[ADD] = '" & RADD & "',  TypeFrame ='R', ti.ItemCode,ti.ItemID,ti.ItemName,ti.SphereID,ti.CylinderID,g.GroupName,tu.UnitName,SPHTable.SPHT,CLYTable.CLYT"
                s = s & "      FROM TblItems AS ti INNER JOIN Groups AS g ON g.GroupID = ti.GroupID"
                s = s & " INNER JOIN TblItemsUnits ON TblItemsUnits.ItemID = ti.ItemID"
                s = s & " INNER JOIN TblUnites AS tu ON tu.UnitID = TblItemsUnits.UnitID"
                s = s & " INNER JOIN SPHTable ON ti.SphereID = SPHTable.ID"
                s = s & " INNER JOIN CLYTable ON ti.CylinderID = CLYTable.ID"
                s = s & " Where  (SPHTable.ID = " & val(GetSphCylID(0, val(RSPH) + val(RADD)))
                s = s & " and CLYTable.ID= " & val(GetSphCylID(1, RCLY)) & ")"
                      
                     If (DCboItemCode.Text = "" And DCboItemName.Text = "") Then
                        s = s & " and TblItems.LensesTypesID  = " & val(cmbGroup1.BoundText)
                        s = s & " and TblItems.GroupID = " & val(cmbGroup.BoundText)
                     End If

                If LADD <> "" Then
                s = s & " Union All                                "
                
                    s = s & " SELECT ItemID2 = " & mMainRow & ",[ADD] = '" & LADD & "', TypeFrame ='L', ti.ItemCode,ti.ItemID,ti.ItemName,ti.SphereID,ti.CylinderID,g.GroupName,tu.UnitName,SPHTable.SPHT,CLYTable.CLYT"
                    s = s & "      FROM TblItems AS ti INNER JOIN Groups AS g ON g.GroupID = ti.GroupID"
                    s = s & " INNER JOIN TblItemsUnits ON TblItemsUnits.ItemID = ti.ItemID"
                    s = s & " INNER JOIN TblUnites AS tu ON tu.UnitID = TblItemsUnits.UnitID"
                    s = s & " INNER JOIN SPHTable ON ti.SphereID = SPHTable.ID"
                    s = s & " INNER JOIN CLYTable ON ti.CylinderID = CLYTable.ID"
                    s = s & " Where  (SPHTable.ID = " & val(GetSphCylID(0, val(LSPH) + val(LADD)))
                    s = s & " and CLYTable.ID= " & val(GetSphCylID(1, LCLY)) & ")"
                                         
                     If (DCboItemCode.Text = "" And DCboItemName.Text = "") Then
                        s = s & " and TblItems.LensesTypesID  = " & val(cmbGroup1.BoundText)
                        s = s & " and TblItems.GroupID = " & val(cmbGroup.BoundText)
                     End If

                End If
                

                loadgrid s, GrdItems2, True, False
                If GrdItems2.rows > 1 Then
                    FramePay.Visible = True
                    GrdItems2.Visible = True
                    GrdItems.left = cmdInsert.Width
                    GrdItems2.left = cmdInsert.Width

                '    GrdItems2.Top =
                    GrdItems.Height = FramePay.Height / 2 - 1000
                    GrdItems.Width = FramePay.Width - 150 - cmdInsert.Width
                    GrdItems2.Width = FramePay.Width - 150 - cmdInsert.Width
                    GrdItems2.Height = FramePay.Height / 2 - 600
                    GrdItems2.top = GrdItems.Height + 1000
                    'Exit Sub

                End If
            Else
                GrdItems2.Visible = False
                GrdItems.Height = FramePay.Height - 200
                GrdItems.Width = FramePay.Width - 150 - cmdInsert.Width
                
            End If
        End If
        
          If Not RSPH Is Nothing Then
           RSPH.Text = ""
           RCLY.Text = ""
           RAxis.Text = ""
           LSPH.Text = ""
           LCLY.Text = ""
           LAxis.Text = ""
           RIPD.Text = ""
           txtRCL = ""
            txtLCL = ""
           
            LIPD.Text = ""
            RADD.Text = ""
            LADD.Text = ""
            RSH.Text = ""
            LSH.Text = ""
            RPRISM.Text = ""
            LPRISM.Text = ""
           
        End If
        mMainRow = 0
    End If
        Else

        If .rows - 2 >= 1 Then
                If .TextMatrix(.rows - 2, .ColIndex("Code")) = "" Then
                    .rows = .rows - 1
                End If
            End If
            
            AddItemInGrid .rows - 1, DCboItemName.BoundText, val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1
            '        customer_screen.bill_no.Caption = frmsalebill.XPTxtBillID.text ' TxtTransSerial.text 'salim2
        
            ' customer_screen.bill_no.Caption = Me.TxtTransSerial.text
            '       customer_screen.List1.AddItem DcboItemName.BoundText & Chr(9) & Val(TxtQuantity.text) & " * " & Me.TxtPrice.text  '& Chr(9) & " = " & Val(TxtQuantity.text) * Me.TxtPrice.text
            '        customer_screen.List2.AddItem Val(TxtQuantity.text) * Me.TxtPrice.text
            '        customer_screen.total
        End If

    End With
    
    If Not RSPH Is Nothing Then
        CheckQty
    End If
    CountItems

    If Me.GridTrans <> MoveItems And Me.GridTrans <> OpeningBalance And Me.GridTrans <> Returntransaction Then
        Calculate Me.Grid.rows - 1
    End If

    If Me.GridTrans = Returntransaction Then
   
        FrmReturnpurchases.FG_AfterEdit 1, 1
     
    End If

    ClearTopDataEntry
     
    If Not TxtCustCode Is Nothing And Me.GridTrans2 = InvoiceTransactionCompose Then
        TxtCustCode.SetFocus
    Else
    
        If txtItemCodeSearch Is Nothing Then
            If DCboItemCode.Enabled Then
            DCboItemCode.SetFocus
            End If
        Else
         If txtItemCodeSearch.Enabled Then
            txtItemCodeSearch.SetFocus
            End If
        End If
        
        
    End If
       
    'Me.DCboItemCode.SetFocus
    'Me.DCboItemCode.SetFocus
    'Me.Grid.SetFocus

    Exit Sub
ErrTrap:
End Sub




Private Function CheckQty() As Boolean
IsOrderOut = False
IsNotSave = True

frmsalebill4.IsOrderOut = IsOrderOut
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim mStoreLab As Long
Dim mStoreLabName As String
Dim s As String
For i = 1 To Grid.rows - 1
    GetItemBalanceInStore i, 1, 1, 1, , val(Grid.TextMatrix(i, Grid.ColIndex("Code"))), DtpBillDate.value
    Grid.TextMatrix(i, Grid.ColIndex("IsOrderOut")) = 0
    Set rsDummy = New ADODB.Recordset
    s = "Select TblStore.* from TblBranchesData"
    s = s & " Inner join TblStore On TblStore.StoreId = TblBranchesData.StoreID"
    s = s & " WHERE BranchId = " & val(Branch.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    
    
    
    
    
    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAvi2")) = ""
    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAviName2")) = ""
    If Not rsDummy.EOF Then
        mStoreLab = val(rsDummy!StoreID & "")
        mStoreLabName = Trim(rsDummy!StoreName & "")
    Else
        mStoreLab = val(StoreName.BoundText)
        mStoreLabName = Trim(StoreName.Text)
    End If
    
    
    s = "SELECT TypeItemsID,IsNull(IsTransfere,0) IsTransfere FROM TblItems AS ti Inner join groups as G On ti.GroupId = G.GroupId  WHERE ti.ItemID =   " & val(Grid.TextMatrix(i, Grid.ColIndex("Code")))
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
    
    
    
    
        Grid.TextMatrix(i, Grid.ColIndex("IsTransfere")) = CBool(rsDummy!IsTransfere & "")
                
        If Not CBool(rsDummy!IsTransfere & "") Then
                
'         If Not CBool(rsDummy!IsTransfere & "") And val(Grid.TextMatrix(i, Grid.ColIndex("ItemBalance"))) < val(Grid.TextMatrix(i, Grid.ColIndex("Count"))) Then
'            CheckQty = False
'            IsNotSave = True
'            frmsalebill4.IsNotSave = IsNotSave
'            Exit Function
            If Not CBool(rsDummy!IsTransfere & "") And val(Grid.TextMatrix(i, Grid.ColIndex("ItemBalance"))) >= val(Grid.TextMatrix(i, Grid.ColIndex("Count"))) Then
                Grid.TextMatrix(i, Grid.ColIndex("IsOrderOut")) = 1
                IsOrderOut = True
                frmsalebill4.IsOrderOut = IsOrderOut
                GoTo NextRow
    '        Else
    '            Grid.TextMatrix(i, Grid.ColIndex("IsOrderOut")) = 1
    '            GoTo NextRow
            End If
        End If
        
       
            
            If val(Grid.TextMatrix(i, Grid.ColIndex("ItemBalance"))) >= val(Grid.TextMatrix(i, Grid.ColIndex("Count"))) Then
                Grid.TextMatrix(i, Grid.ColIndex("RequestTypeName")) = "ÿ·»  ÕÊÌ· „Œ“‰Ì"
                Grid.TextMatrix(i, Grid.ColIndex("RequestTypeNo")) = "2"
                

                Grid.TextMatrix(i, Grid.ColIndex("StoreIDAvi2")) = mStoreLab
                Grid.TextMatrix(i, Grid.ColIndex("StoreIDAviName2")) = mStoreLabName
            ElseIf val(Grid.TextMatrix(i, Grid.ColIndex("ItemBalance"))) < val(Grid.TextMatrix(i, Grid.ColIndex("Count"))) Then
                If val(Grid.TextMatrix(i, Grid.ColIndex("RequestTypeNo"))) = 1 Then
                    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAvi2")) = mStoreLab
                    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAviName2")) = mStoreLabName
                ElseIf val(Grid.TextMatrix(i, Grid.ColIndex("RequestTypeNo"))) = 2 Then
                    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAvi2")) = mStoreLab
                    Grid.TextMatrix(i, Grid.ColIndex("StoreIDAviName2")) = mStoreLabName
                End If
                
            End If
            
       
        
    End If
NextRow:
Next
CheckQty = True

End Function

Private Sub cmdAdd_Click()


If FramePay.Visible = False Then
Else
    
End If
End Sub

Private Function GetSphCylID(ByVal mType As Integer, ByVal MSTR As String) As Long

Dim s As String
If mType = 0 Then
    s = "Select Id from SPHTable Where SPH =  " & val(MSTR)
ElseIf mType = 1 Then
    s = "Select Id from CLYTable Where CLY  =  " & val(MSTR)
End If
Dim rsDummy As ADODB.Recordset
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    GetSphCylID = val(rsDummy!ID & "")
End If
End Function

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
    Dim StrSQL    As String
    Dim RsTemp    As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
    TxtPrice.Text = 0
    If Me.GridTrans <> GridTransType.purchaseOrderApproved Then
        '›Ï Õ«·Â «„— «·‘—«¡ Ì „ Ÿ»ÿ «·ﬂ„ÌÂ „‰ ⁄—Ê÷ «·«”⁄«—
        '                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
        '                        Else
        '                            purQty = 1 ' defualt
        '                        End If
        TxtQuantity.Text = ""
    End If

    '   If DCboItemCode.BoundText <> "" Then
    '       DcboItemName.BoundText = DCboItemCode.BoundText
    '   Else
    '       Exit Sub
    '   End If

    StrSQL = "select * From TblItems where ItemID=" & val(DCboItemCode.BoundText)
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
              
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                If Me.GridTrans = GridTransType.purchaseOrderApproved Then
                    '›Ï Õ«·Â «„— «·‘—«¡ Ì „ Ÿ»ÿ «·ﬂ„ÌÂ „‰ ⁄—Ê÷ «·«”⁄«—
                    '                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
                    '                        Else
                    '                            purQty = 1 ' defualt
                    '                        End If
                    If val(TxtQuantity.Text) = 0 Then
 
                        TxtQuantity.Text = 1
                    End If
                    
                Else
                    TxtQuantity.Text = 1
                End If
            End If
            TxtQuantity.Tag = "T"

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = 3 Then
         
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
                 
                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
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
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
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
                
                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                End If
                If val(TxtPrice.Text) = 0 Then
                    TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
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
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
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

        For RowNum = 1 To .rows - 1

            If .TextMatrix(RowNum, .ColIndex("Code")) <> "" Then

                ' FG_AfterEdit RowNum, Fg.ColIndex("Price")
                'Grid_AfterEdit RowNum, .ColIndex("Code")
                If m_GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Or m_GridTrans = Destruction Or m_GridTrans = 37 Then
                    LngItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
             '       LngUnitID = val(.TextMatrix(RowNum, .ColIndex("UnitId")))
    LngUnitID = val(Grid.cell(flexcpData, RowNum, Grid.ColIndex("UnitID")))
   Dim transferid As Double
   
    If m_GridTrans = MoveItems Then
    
  transferid = Get_movingreciveTransaction_ID(val(TxtInvID.Text))
  If LngItemID = 2249 Then
  LngItemID = 2249
  End If
  .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(transferid), LngUnitID, val(Me.StoreName.BoundText))
  Else
 
  
  .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))

    End If
    
                    
       
       
                End If
                    
                Calculate (RowNum), True
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
Public Sub DcboItemCodeFromTextSerailCode()
If txtItemCodeSearch2.Text <> "" Then
DCboItemCode_KeyDown vbKeyReturn, 0
End If
txtItemCodeSearch2.Text = ""
txtItemCodeSearch2.SetFocus
End Sub

Public Sub DCboItemCode_KeyDown(KeyCode As Integer, _
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
      If Me.GridTrans = InvoiceTransaction Then
                                               TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
                                               
   End If
                                            
                                                If val(DCboItemCode.BoundText) <> 0 And FrmCustomerDisplay.Visible = True Then
                                               
                                                
                                               
                                                
                                            '''''FrmCustomerDisplay.LblInformation.Caption = DcboItemName.text & vbNewLine & "  Price   " & val(Me.TxtPrice.text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
                                           
                                               FrmCustomerDisplay.lblitemid = val(DCboItemCode.BoundText)
                                          

                                              End If
                                             
                                               
                                    Sendkeys "{TAB}"
                                    Sendkeys "{TAB}"
                                    Sendkeys "{TAB}"
                                          
                                                      
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                TxtQuantity.Text = 1
            End If

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
   
   
        Sendkeys "{TAB}"

        
    End If

End Sub

Private Sub DCboItemCode_LostFocus()
    DCboItemCode.backcolor = vbWhite
    'TxtPrice.BackColor = vbWhite
End Sub

'
Private Sub DcboItemName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL    As String
    Dim RsTemp    As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
    TxtPrice.Text = 0
    'TxtQuantity.text = ""
    If Me.GridTrans <> GridTransType.purchaseOrderApproved Then
        '›Ï Õ«·Â «„— «·‘—«¡ Ì „ Ÿ»ÿ «·ﬂ„ÌÂ „‰ ⁄—Ê÷ «·«”⁄«—
        '                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
        '                        Else
        '                            purQty = 1 ' defualt
        '                        End If
        TxtQuantity.Text = ""
    End If
    
    If DCboItemName.BoundText <> "" Then
        DCboItemCode.BoundText = DCboItemName.BoundText
      
        If Me.GridTrans = InvoiceTransaction And Not (txtItemCodeSearch2 Is Nothing) Then
            txtItemCodeSearch2.Text = DCboItemCode.Text
        End If
        
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
                        
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
                If Me.GridTrans = GridTransType.purchaseOrderApproved Then
                    '›Ï Õ«·Â «„— «·‘—«¡ Ì „ Ÿ»ÿ «·ﬂ„ÌÂ „‰ ⁄—Ê÷ «·«”⁄«—
                    '                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
                    '                        Else
                    '                            purQty = 1 ' defualt
                    '                        End If
                    If val(TxtQuantity.Text) = 0 Then
                        TxtQuantity.Text = 1
                    End If
                Else
                    TxtQuantity.Text = 1
                End If
            
            End If

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
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
            'SalimSalim
              
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = 0
            Else
            If Me.GridTrans = GridTransType.purchaseOrderApproved Then
                    '›Ï Õ«·Â «„— «·‘—«¡ Ì „ Ÿ»ÿ «·ﬂ„ÌÂ „‰ ⁄—Ê÷ «·«”⁄«—
                    '                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
                    '                        Else
                    '                            purQty = 1 ' defualt
                    '                        End If
                    If val(TxtQuantity.Text) = 0 Then
                        TxtQuantity.Text = IIf(SystemOptions.DefaultQtyTrans = 0, "", SystemOptions.DefaultQtyTrans)
                    End If
                Else
                    TxtQuantity.Text = IIf(SystemOptions.DefaultQtyTrans = 0, "", SystemOptions.DefaultQtyTrans)
                End If
               
            End If

            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Then
                TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
                
                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                    If TxtPrice.Text = 0 Then
                        TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
                    End If
                        
                    Dim rsDummy As New ADODB.Recordset
                        
                    '     StrSQL = "Select Emp_ID,CustomerId,FromDate,Todate,TblCustomerContractDetails.Price  from TblCustomerContract"
                    '     StrSQL = StrSQL & " Left Outer join TblCustomerContractDetails On TblCustomerContractDetails.TblCustomerContractD = TblCustomerContract.TblCustomerContractD"
                    '     StrSQL = StrSQL & " Where customerid = " & val(Customer.BoundText)
                    '     StrSQL = StrSQL & " and TblCustomerContractDetails.ItemID =" & val(DCboItemName.BoundText)
                    '
                    '     StrSQL = StrSQL & " and " & SQLDate(DtpBillDate.value, True) & "   BETWEEN FromDate and Todate"
                    '     Set rsDummy = New ADODB.Recordset
                    '     rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    '     If Not rsDummy.EOF Then
                    '         TxtPrice.Enabled = False
                    '         TxtPrice = rsDummy!Price & ""
                    '
                    '     Else
                    '         TxtPrice.Enabled = True
                    '                        End If
                        
                    '
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
            
                FromDate = FirstPeriodDateInthisYear
                GetlastPurchasedata 20, (LngItemID), FromDate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                
                TxtPrice.Text = LastPurchasePrice
                   
            Else
                TxtPrice.Text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
          
                If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                    TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
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
   
           Sendkeys "{TAB}"
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

    Select Case Button.key

        Case "AddRow"

            If Me.Grid.TextMatrix(Me.Grid.rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
                Me.Grid.rows = Me.Grid.rows + 1
                Me.GridDefaultValue Me.Grid.rows - 1
                Me.Grid.Row = Me.Grid.rows - 1
                Me.Grid.Col = Me.Grid.ColIndex("Code")
                Me.Grid.ShowCell Me.Grid.rows - 1, Me.Grid.ColIndex("Code")
                Me.Grid.SetFocus
            End If
 Case "AddRow2"

            If Me.Grid.TextMatrix(Me.Grid.rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
                Grid.AddItem " ", Grid.Row + 1
                Me.GridDefaultValue Me.Grid.Row + 1
                Grid.Row = Grid.Row + 1
                Me.Grid.Col = Me.Grid.ColIndex("Code")
                Me.Grid.ShowCell Me.Grid.Row + 1, Me.Grid.ColIndex("Code")
                Me.Grid.SetFocus
            End If

        Case "RemoveRow"
        Dim F As Form
            For Each F In Forms
        If F.Name = "frmsalebill1" Then
      '      MsgBox "Form Loaded"
            Exit Sub
        End If
        
        
        
                If F.Name = "frmsalebill2" Then
                                      If SystemOptions.HideTbarInPos = True Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                     MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— €Ì— „”„Ã ·ﬂ"
                                                Else
                                                MsgBox "    not authorized"
                                                End If
                                                  Exit Sub
                                End If
                                   
        End If

     
     
    Next F
    
   
            Dim IsSelect As Boolean
            Dim i As Long
            If Grid.ColIndex("Select") <> -1 Then
                For i = 1 To Grid.rows - 1
                    If i < Grid.rows Then
                        If val(Me.Grid.TextMatrix(i, Me.Grid.ColIndex("Select"))) <> 0 Then
                            Me.Grid.RemoveItem i
                            IsSelect = True
                            i = i - 1
                        End If
                    End If
                Next
                SentTypeVAT
                    Calculate 1, , , True
                    RelinVatGrid
            End If
            If IsSelect Then Exit Sub
            If Grid.rows > 1 Then
                If Grid.rows = 2 Then
                    Me.Grid.Clear flexClearScrollable, flexClearEverything
                    Me.Calculate 1, True, , True
                Else

                    If Me.Grid.rows > 1 Then
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
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

            If IntRes = vbYes Then
                
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Grid.rows = Me.Grid.FixedRows + 1
                Me.Grid.Row = 1
                Me.VatGrid.Clear flexClearScrollable, flexClearEverything
                Me.VatGrid.rows = Me.Grid.FixedRows + 1
                Me.VatGrid.Row = 1
                
            End If
RelinVatGrid
        Case "FillData"

            If Me.GridTrans = OpeningBalance Then
                If Grid.TextMatrix(Grid.rows - 1, Grid.ColIndex("Code")) <> "" Then
                    Grid.rows = Grid.rows + 1
                    Me.GridDefaultValue Me.Grid.rows - 1
                End If

          '      FrmFillBalance.show vbModal
            Else

                If Me.Grid.TextMatrix(Me.Grid.rows - 1, Me.Grid.ColIndex("Code")) <> "" Then
                    Me.Grid.rows = Me.Grid.rows + 1
                    Me.GridDefaultValue Me.Grid.rows - 1
                End If

             '   FrmFillItems.DealingForm = Me.GridTrans
             '   FrmFillItems.show vbModal
             '   Calculate 1, , , True
            End If

        Case "Tree"
If SystemOptions.OPenShortInvoice = True Then

Exit Sub

End If


            If Button.value = tbrPressed Then
                Set MyAddItemForm = New FrmAddItem
                Set MyAddItemForm.cGrid = Me
                MyAddItemForm.show
            Else
                CloseMyAddItemForm
            End If

        Case "Barcode"
If SystemOptions.OPenShortInvoice = True Then

Exit Sub

End If

            If Me.Grid.Parent.XPTxtBillID.Text <> "" Then
            
                Load FrmPrintBarcode
                FrmPrintBarcode.LblID.Caption = Me.Grid.Parent.XPTxtBillID.Text
         If Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Then
 FrmPrintBarcode.DBCboClientName.BoundText = Me.Customer.BoundText
  FrmPrintBarcode.DBCboClientName_Click (0)
         End If
         
         
                
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
Function GetMaxUnitFactors(Optional ItemID As Double, Optional UnitID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(UnitFactor) AS MaxUnitFactor"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (UnitFactor <> 1 and ItemID = " & ItemID & ")"
'sql = sql & " and  (unitid = " & UnitID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxUnitFactors = IIf(IsNull(rs2("MaxUnitFactor").value), 0, rs2("MaxUnitFactor").value)
Else
GetMaxUnitFactors = 0
End If
End Function

Function GetUnitFactors(Optional ItemID As Double, Optional UnitID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     MAX(UnitFactor) AS UnitFactors"
sql = sql & " From dbo.TblItemsUnits"
sql = sql & " Where (ItemID = " & ItemID & ")"
sql = sql & " and  (unitid = " & UnitID & ")"
rs2.Open sql, Cn, adOpenForwardOnly, adLockReadOnly
If rs2.RecordCount > 0 Then
GetUnitFactors = IIf(IsNull(rs2("UnitFactors").value), 0, rs2("UnitFactors").value)
Else
GetUnitFactors = 0
End If
End Function

Function GetMaxUnitID(Optional ItemID As Double, Optional UnitFactor As Double, Optional mUnitId As Long = 0) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT       UnitID  "
sql = sql & " From dbo.TblItemsUnits"
'«·ÿ—Ìﬁ… «·ﬁœÌ„… ﬂ«‰ »ÌÃÌ» «ﬂ»— ﬁÌ„… ··„⁄«„· ÊÌ⁄ »— «‰Â «·ÊÕœ… «·ﬂ»—Ï
'sql = sql & " Where (ItemID = " & ItemID & ") and UnitFactor=" & UnitFactor & ""
'«·ÿ—Ìﬁ… «·«ŒÌ—…  »ÌÃÌ» «·ÊÕœ… «·«Œ—Ï
'sql = sql & " Where (ItemID = " & ItemID & ") and UnitID<> " & mUnitId & "  order by UnitFactor desc"
sql = sql & " Where (ItemID = " & ItemID & ")  order by UnitFactor desc"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxUnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
Else
GetMaxUnitID = 0
End If
End Function
 Public Sub CalculteMaxQtyUnit(LongRow As Long, Optional ByVal IsMaxQty As Boolean = False)
Dim UnitFactor As Double
Dim MaxUnitFactor As Double
Dim LngItemID As Double
Dim MaxUnitID As Double
Dim UnitID As Long
Dim LngUnitID  As Long
Dim DivFactor As Double
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))

'UnitID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("UnitID")))
 LngUnitID = val(Grid.cell(flexcpData, LongRow, Grid.ColIndex("UnitID")))
UnitFactor = GetUnitFactors(LngItemID, CDbl(LngUnitID))
If val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxUnitID"))) = 0 Then
 Grid.TextMatrix(LongRow, Grid.ColIndex("MaxUnitID")) = GetMaxUnitID(LngItemID, UnitFactor, LngUnitID)
End If
MaxUnitID = Grid.TextMatrix(LongRow, Grid.ColIndex("MaxUnitID"))
MaxUnitFactor = GetUnitFactors(LngItemID, MaxUnitID)


If UnitFactor > MaxUnitFactor Then
DivFactor = UnitFactor / MaxUnitFactor
Else
DivFactor = MaxUnitFactor / IIf(UnitFactor <> 0, UnitFactor, 1)
End If
 

    
  
    

    If Not IsMaxQty Then
        If MaxUnitFactor <> 0 Then
        
         
'                If UnitFactor = 1 Then
'                Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Count")))) / MaxUnitFactor, 3)
'                Else
'                 Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Count")))) * MaxUnitFactor, 3)
'                End If
  
    If UnitFactor > MaxUnitFactor Then 'div
    Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Count")))) * DivFactor, 3)
    Else
     Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Count")))) / DivFactor, 3)
    End If
      
      

        Else
            Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")) = 0
        End If
        
        
    Else
        If val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty"))) <> 0 Then
         
           ' Grid.TextMatrix(LongRow, Grid.ColIndex("Count")) = val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty"))) * MaxUnitFactor
           ' If UnitFactor < MaxUnitFactor Then
           ' Grid.TextMatrix(LongRow, Grid.ColIndex("Count")) = val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty"))) / UnitFactor
           ' End If
                   If MaxUnitFactor <> 0 Then
        
         
                If UnitFactor > MaxUnitFactor Then 'div
                Grid.TextMatrix(LongRow, Grid.ColIndex("Count")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")))) / DivFactor, 3)
                Else
                 Grid.TextMatrix(LongRow, Grid.ColIndex("Count")) = Round((val(Grid.TextMatrix(LongRow, Grid.ColIndex("MaxQty")))) * DivFactor, 3)
                End If
  

        Else
            Grid.TextMatrix(LongRow, Grid.ColIndex("Count")) = 0
        End If
            
            
        End If
    End If
End Sub

Public Sub Grid_AfterEdit(ByVal Row As Long, _
   ByVal Col As Long)
 
    Dim RowNum         As Integer
    Dim SearchResault  As Integer
    Dim RsSerial       As ADODB.Recordset
    Dim RsFind         As ADODB.Recordset
    Dim RsTemp         As ADODB.Recordset
    Dim RsTest         As ADODB.Recordset
    Dim RsPrice        As ADODB.Recordset
    Dim StrSQL         As String
    Dim Msg            As String
    Dim LngItemID      As Long
    Dim LngUnitID      As Long
    Dim StrAccountCode As String
    Dim ClsAcc As New ClsAccounts
    Dim rsDummy As New ADODB.Recordset
    Dim rsDummyPrice As New ADODB.Recordset
    Dim mPrice As Double
    On Error GoTo ErrTrap
    Dim LngRow As Long
    Dim mTotal As Long
    If TxtFillData.Text = "T" Then Exit Sub
    If Row = -1 Or Col = -1 Then Exit Sub
    If RsItems Is Nothing Then Exit Sub
    If SystemOptions.PriceWithVAT = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) Then VatGrid.rows = 1: If Not LblValueAdded Is Nothing Then LblValueAdded.Caption = ""
    With Me.Grid

        Select Case .ColKey(Col)
            Case "ItemCodeSearch"
                Dim s       As String
                
                s = "Select * from tblItems Where Code = N'" & Trim(Grid.TextMatrix(Row, Grid.ColIndex("ItemCodeSearch"))) & "'"
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                        
                    Grid.TextMatrix(Row, Grid.ColIndex("Code")) = val(rsDummy!ItemID & "")
                    Grid_AfterEdit Row, Grid.ColIndex("Code")
                    
                End If
            Case "CurrentQty"
                .TextMatrix(Row, .ColIndex("Count")) = val(.TextMatrix(Row, .ColIndex("CurrentQty"))) - val(.TextMatrix(Row, .ColIndex("PrevQty")))
                .TextMatrix(Row, .ColIndex("StillPumbQty")) = .TextMatrix(Row, .ColIndex("Count"))
'                .TextMatrix(LngRow, .ColIndex("PrevQty")) = txtPumbDet1
'                .TextMatrix(LngRow, .ColIndex("currentqty")) = txtPumbDet2
            Case "Price"
                mPrice = Grid.TextMatrix(Row, Grid.ColIndex("Price"))
                If Me.GridTrans = purchaseorder Or Me.GridTrans = purchaseorderrequest Or Me.GridTrans = PurchaseTransaction Then
                    Exit Sub
                End If

                If SystemOptions.CanChangePriceUpOnly Then
                 Set rsDummyPrice = New ADODB.Recordset
                                    
                    StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID,TBLiTEMSUnits.UnitSalesPrice, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                                    
                    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                    StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId "
                    'and DefaultUnit = 1"
                    StrSQL = StrSQL + " Where TblItems.ItemID=" & val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))
                    StrSQL = StrSQL + "  and TBLiTEMSUnits.UnitID=" & val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                    rsDummyPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummyPrice.EOF Then
                        If mPrice < val(rsDummyPrice!UnitSalesPrice & "") Then
                             Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(rsDummyPrice!UnitSalesPrice & "")
                        End If
                    End If
                    
                End If
            Case "MaxQty"
                CalculteMaxQtyUnit Row, True
            Case "MaxUnitID"
                CalculteMaxQtyUnit Row, True
                Case "CashQty"
                  ', "MadaQty", "VisaQty", "DeferredQty"
                    .TextMatrix(Row, .ColIndex("Cash")) = val(.TextMatrix(Row, .ColIndex("CashQty"))) * val(.TextMatrix(Row, .ColIndex("Price"))) '* 1.15
                    
                     mTotal = val(.TextMatrix(Row, .ColIndex("CashQty"))) + val(.TextMatrix(Row, .ColIndex("MadaQty"))) + val(.TextMatrix(Row, .ColIndex("VisaQty"))) + val(.TextMatrix(Row, .ColIndex("DeferredQty")))
                    .TextMatrix(Row, .ColIndex("StillPumbQty")) = val(.TextMatrix(Row, .ColIndex("Count"))) - mTotal
                  Case "MadaQty"
                  ', "MadaQty", "VisaQty", "DeferredQty"
                    .TextMatrix(Row, .ColIndex("Mada")) = val(.TextMatrix(Row, .ColIndex("MadaQty"))) * val(.TextMatrix(Row, .ColIndex("Price"))) '* 1.15
                    mTotal = val(.TextMatrix(Row, .ColIndex("CashQty"))) + val(.TextMatrix(Row, .ColIndex("MadaQty"))) + val(.TextMatrix(Row, .ColIndex("VisaQty"))) + val(.TextMatrix(Row, .ColIndex("DeferredQty")))
                    .TextMatrix(Row, .ColIndex("StillPumbQty")) = val(.TextMatrix(Row, .ColIndex("Count"))) - mTotal

                    
                  Case "VisaQty"
                  ', "MadaQty", "VisaQty", "DeferredQty"
                    .TextMatrix(Row, .ColIndex("Visa")) = val(.TextMatrix(Row, .ColIndex("VisaQty"))) * val(.TextMatrix(Row, .ColIndex("Price"))) '* 1.15
                                         mTotal = val(.TextMatrix(Row, .ColIndex("CashQty"))) + val(.TextMatrix(Row, .ColIndex("MadaQty"))) + val(.TextMatrix(Row, .ColIndex("VisaQty"))) + val(.TextMatrix(Row, .ColIndex("DeferredQty")))
                    .TextMatrix(Row, .ColIndex("StillPumbQty")) = val(.TextMatrix(Row, .ColIndex("Count"))) - mTotal

            Case "RequestTypeName"
                Grid.TextMatrix(Row, Grid.ColIndex("RequestTypeNo")) = val(Grid.TextMatrix(Row, Grid.ColIndex("RequestTypeName")))
        
        
        
                 Case "Account_Serial2"
            

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Code='" & Trim(.TextMatrix(Row, Col)) & "'"
                StrSQL = StrSQL & GetAccountByBarnchUser
                Set rsDummy = Nothing
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rsDummy.BOF Or rsDummy.EOF) Then
                   

                    .TextMatrix(Row, .ColIndex("Account_Code")) = IIf(IsNull(rsDummy("Account_Code").value), "", rsDummy("Account_Code").value)
                    .TextMatrix(Row, .ColIndex("Account_Serial2")) = IIf(IsNull(rsDummy("Account_Serial").value), "", rsDummy("Account_Serial").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("Account_Name")) = IIf(IsNull(rsDummy("Account_Name").value), "", rsDummy("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("Account_Name")) = IIf(IsNull(rsDummy("Account_NameEng").value), "", rsDummy("Account_NameEng").value)
                    
                    End If
                    
xx:
                Else
                    'GetMsgs 130, vbExclamation
                    If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "ﬂÊœ Õ”«» Œ«ÿÏ¡", vbCritical
                    Else
                    MsgBox "Error Code Account", vbCritical
                    End If
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("Account_Code")) = ""
                    
                
                    Exit Sub
                End If

                rsDummy.Close
                Set rsDummy = Nothing

            Case "Account_Name"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                
                      
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Account_Name"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                    .TextMatrix(Row, .ColIndex("Account_Code")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial2")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    .TextMatrix(Row, .ColIndex("Account_Name")) = ClsAcc.Get_Account_Name(StrAccountCode)

                
                Set ClsAcc = Nothing
            
               
                
'
'            Case "Account_Name"
'                StrAccountCode = Grid.ComboData
'
'                LngRow = .FindRow(StrAccountCode, Grid.FixedRows, Grid.ColIndex("Account_Name"), False, True)
'                Grid.TextMatrix(Row, Grid.ColIndex("Account_Code")) = StrAccountCode
            Case "Account_Serial2"
                StrAccountCode = Grid.ComboData
            
                LngRow = .FindRow(StrAccountCode, Grid.FixedRows, Grid.ColIndex("Account_Serial2"), False, True)
                Grid.TextMatrix(Row, Grid.ColIndex("Account_Code")) = StrAccountCode

            Case "StoreIDAviName"
                StrAccountCode = Grid.ComboData
            
                LngRow = .FindRow(StrAccountCode, Grid.FixedRows, Grid.ColIndex("StoreIDAviName"), False, True)
                Grid.TextMatrix(Row, Grid.ColIndex("StoreIDAvi")) = StrAccountCode
                
                GetBalnceOfItems val(StrAccountCode), val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), Row
                
             Case "Code"

                If Grid.TextMatrix(Row, Col) <> "" Then
                    LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, adBookmarkFirst

                    If RsItems.EOF Or RsItems.BOF Then
                        Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Grid.TextMatrix(Row, Col) = ""
                        Me.Grid.Row = Row
                        Me.Grid.Col = Col
                        Exit Sub
                    Else
                        Dim purQty As Double
                        If Me.GridTrans = GridTransType.purchaseOrderApproved Then
                            purQty = val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) ' √„— «·‘—«¡
                        Else
                            purQty = 1 ' defualt
                        End If
                     
                        AddItemInGrid Row, LngItemID, purQty
                        LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                        LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
               
                        LngUnitID = val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                        '  .TextMatrix(Row, .ColIndex("StoreIDLoc")) = GetLocationID(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))))
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
                            .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                        End If
            
                        '*************************
                        If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                            .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                            If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                                .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID)
                            End If
                        End If
                
                        '**************
                    End If
                   
                    If Grid.ColIndex("RequestTypeName") = -1 Then
                        Exit Sub
                    Else
                        If Me.GridTrans <> GridTransType.purchaseOrderApproved Then
                            GetBalnceOfItems val(Me.StoreName.BoundText), val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), Row
                        End If
                    End If

                Else
                    Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                    'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Grid.cell(flexcpText, Row, Col, Row, .Cols - 1) = ""
                    Me.Grid.Row = Row
                    Me.Grid.Col = Col
                End If
                   
            Case "Name"

                If Grid.TextMatrix(Row, Col) <> "" Then
                    LngItemID = val(Grid.ComboData)
                    RsItems.Find "ItemID=" & LngItemID, , adSearchForward, 1

                    If RsItems.EOF Or RsItems.BOF Then
                        Msg = "·« ÊÃœ »Ì«‰«  ⁄‰ Â–« «·’‰›" & CHR(13)
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Grid.TextMatrix(Row, Col) = ""
                        Me.Grid.Row = Row
                        Me.Grid.Col = Col
                        Exit Sub
                    Else
                        '   If Me.DtpBillDate.value <> Null Then
                        
                        AddItemInGrid Row, LngItemID, , , , , , , , , , , , , , , Me.DtpBillDate.value
                       
                        LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                        LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
               
                        LngUnitID = val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                        '   .TextMatrix(Row, .ColIndex("StoreIDLoc")) = GetLocationID(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))))
                        '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                        If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Or m_GridTrans = Returntransaction Or m_GridTrans = 28 Then
                            AddItemPrice LngItemID, Col, Row, LngUnitID
                        ElseIf m_GridTrans = PurchaseTransaction Then
                            AddItemPrice LngItemID, Col, Row, LngUnitID, 1
                           
                        End If
                           
                        If SystemOptions.workWithvendorContract = True And (Me.GridTrans = PurchaseTransaction Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = 24) Then
                            .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, 1, val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                        End If
                
                        '*************************
                        If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                            .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                            If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                                .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
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
                Dim rdate      As Date
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
                        
            Case "Width", "Height", "NoCount", "length", "OUTR", "INR", "thickness"
            
                'SalimHereAnd Wael
                checkQtybylengthoption Row, LngItemID
                'If Codition
              
            Case "Area"
                Grid.TextMatrix(Row, Grid.ColIndex("Width")) = 0
                Grid.TextMatrix(Row, Grid.ColIndex("Height")) = 0
                Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
            Case "NoCount"
                Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
           
            Case "Serial"
            
                LngItemID = val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                SearchResault = Grid.FindRow(Grid.TextMatrix(Row, Grid.ColIndex("Serial")), , Grid.ColIndex("Serial"))

                If Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then
                    '›Ï Õ«·… ‘«‘… «· ·›Ì«  ‰Õ‰ ‰÷⁄ ”⁄— «· ﬂ·›…
                    '      StrSerial = Trim$(.TextMatrix(Row, .ColIndex("Serial")))
                    Dim UnitID As Long
 
                    UnitID = val(.cell(flexcpData, Row, .ColIndex("UnitID")))
                
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
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
                                                Msg = Msg + Trim(.TextMatrix(Row, .ColIndex("Serial"))) & CHR(13) & "„‰ «·’‰› : " & Trim(.cell(flexcpTextDisplay, Row, .ColIndex("Name"))) & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  " & Row
                                                Msg = Msg + " €Ì— „ÊÃÊœ… ›Ì «·„Œ“‰ «·„Õœœ" & CHR(13)
                                                Msg = Msg + "Â·  —€»  ›Ì „⁄—›… ”Ì—Ì«· «·ﬁÿ⁄ «·„ÊÃÊœ… ›Ï «·„Œ«“‰"
                                
                                                If (MsgBox(Msg, vbYesNo + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)) = vbYes Then
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

            Case "project"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("projectid"), False, True)
                .TextMatrix(Row, .ColIndex("projectid")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("operaid")) = 0
                .TextMatrix(Row, .ColIndex("pandid")) = 0
                .TextMatrix(Row, .ColIndex("pand")) = ""
                .TextMatrix(Row, .ColIndex("opera")) = ""
               Case "pand"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("pandid"), False, True)
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                 .TextMatrix(Row, .ColIndex("operaid")) = 0
                .TextMatrix(Row, .ColIndex("opera")) = ""
              Case "opera"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("operaid"), False, True)
                .TextMatrix(Row, .ColIndex("operaid")) = StrAccountCode

            Case "guaranteeTime" ' «·÷„«‰

                If .TextMatrix(Row, .ColIndex("guaranteeTime")) <> "" Then
                    If Not (IsNumeric(.TextMatrix(Row, .ColIndex("guaranteeTime")))) Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) < 1 Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) > 36 Or val(.TextMatrix(Row, .ColIndex("guaranteeTime"))) <> Int(val(.TextMatrix(Row, .ColIndex("guaranteeTime")))) Then
                        Msg = "Â–Â «·ﬁÌ„… €Ì— ’«·Õ…" & CHR(13)
                        Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì… ’ÕÌÕ… »Ì‰ 1 -36 " & CHR(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    End If
                End If

            Case "Count"

                '  If mdifrmmain.ActiveForm.name = "FrmReturnSalling" Or mdifrmmain.ActiveForm.name = "FrmReturnpurchases" Then
                If m_GridTrans = ReturnSalling Or m_GridTrans = Returntransaction Then
                    If m_GridTrans = ReturnSalling Then
                        StrSQL = "   SELECT     SUM(ShowQty) AS ShowQty FROM         dbo.Transaction_Details   Where Transaction_ID=" & val(TxtInvID.Text) & " and item_id=" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                        '
  
                    Else
                    
                        StrSQL = " SELECT     SUM(ShowQty) AS ShowQty  FROM         dbo.Transaction_Details where  Transaction_ID=" & val(TxtInvID.Text) & " and item_id=" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
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
                If val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) <> 4 Then
             
                    ShowPricess Row, False, val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))), LngItemID
                End If
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
                If LngUnitID = 0 Then
                    GetDefaultItemUnit LngItemID, LngUnitID
                End If
                   
                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Then
                    If SystemOptions.IsGeometricProportions Then
                    
                        LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitId")))
               
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
                   
                Me.Grid.cell(flexcpData, Row, Col) = Grid.ComboData

                '≈÷«›… ”⁄— «·»Ì⁄ «·Œ«’ »«·’‰›
                If m_GridTrans = InvoiceTransaction Or m_GridTrans = ShowPrice Or m_GridTrans = 28 Then
                    AddItemPrice LngItemID, Col, Row, LngUnitID
                
                    Me.Grid.cell(flexcpData, Row, Col) = Grid.ComboData
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
                Me.Grid.cell(flexcpData, Row, Col) = Grid.ComboData
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
            If val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 Then
                Grid.TextMatrix(Row, Grid.ColIndex("Price")) = 0
            End If
        End If
        If val(Grid.TextMatrix(Row, Grid.ColIndex("Price"))) = 0 Then
            Grid.TextMatrix(Row, Grid.ColIndex("Price")) = GetItemPrice(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), , val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID"))))
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
        
            If .rows = .FixedRows + 1 Then
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

        Grid.TextMatrix(Row, Grid.ColIndex("SalesValue")) = GetItemPrice(val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), , val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID"))))
        If Grid.ColKey(Col) <> "MaxQty" Then

            CalculteMaxQtyUnit Row
        End If
   
    End With
    
    If Not RSH Is Nothing Then
        CheckQty
    End If
   ReCalVAT Row, 1
    If Me.GridTrans = PurchaseTransaction Then
        FrmBillBuy.chkTaxExempt_Click
        
    End If
    If Me.GridTrans = InvoiceTransaction And UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
        frmsalebill6.CalcValueGrids
    End If
    If Me.GridTrans = InvoiceTransaction And UCase(frmname) = "FRMSALEBILL" Then
        frmsalebill.chkTaxExempt_Click
        
    End If
    'salman TxtDiscount_Val_Change
    Exit Sub
ErrTrap:
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
               Dim LngUnitID  As Integer
                
                 Grid.TextMatrix(Row, Grid.ColIndex("Area")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("thickness"))) * val(.TextMatrix(Row, .ColIndex("Height")))
                     LngItemID = val(.TextMatrix(Row, .ColIndex("Name")))
                     
                   '  GetItemPriceByWitdth2
                   If SystemOptions.AllowPriceWithWidth = True Then
                            LngUnitID = val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                      '      LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitID")))
                              widthPrice = GetItemPriceByWitdth(LngItemID, val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))), LngUnitID)
                             ' Grid.TextMatrix(Row, Grid.ColIndex("Area")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("thickness")))
                             
                             Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
'                             Grid.TextMatrix(Row, Grid.ColIndex("ShowQty")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
                              .TextMatrix(Row, .ColIndex("Price")) = widthPrice '* val(Grid.TextMatrix(Row, Grid.ColIndex("Area"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount")))
                              Grid_AfterEdit Row, Grid.ColIndex("Count")
                              '.TextMatrix(Row, .ColIndex("Price")) = widthPrice * val(Grid.TextMatrix(Row, Grid.ColIndex("Area"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount")))
                   Else
                           
                          If mIsCountIsGeometricProportions Then
                               Grid.TextMatrix(Row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
                            Else
                                LngUnitID = val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                                 
                          '      LngUnitID = val(Grid.TextMatrix(Row, Grid.ColIndex("UnitID")))
                                  'widthPrice = GetItemPriceByWitdth(LngItemID, val(Grid.TextMatrix(row, Grid.ColIndex("Width"))), LngUnitID)
                                 ' Grid.TextMatrix(Row, Grid.ColIndex("Area")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("thickness")))
                                 
                                 'Grid.TextMatrix(row, Grid.ColIndex("Count")) = val(Grid.TextMatrix(row, Grid.ColIndex("NoCount"))) * val(Grid.TextMatrix(row, Grid.ColIndex("Area")))
    '                             Grid.TextMatrix(Row, Grid.ColIndex("ShowQty")) = val(Grid.TextMatrix(Row, Grid.ColIndex("Area")))
                                  .TextMatrix(Row, .ColIndex("Price")) = GetItemPriceByWitdth2(LngItemID, CInt(LngUnitID), val(Grid.TextMatrix(Row, Grid.ColIndex("Width"))), val(Grid.TextMatrix(Row, Grid.ColIndex("length"))), val(Grid.TextMatrix(Row, Grid.ColIndex("Height"))))   '* val(Grid.TextMatrix(Row, Grid.ColIndex("Area"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("NoCount")))
                                  Grid_AfterEdit Row, Grid.ColIndex("Count")

                                
                           End If
                   End If
                   
                   
                   
            End With
             
End Sub

 
Private Function GetItemPriceByWitdth2(Item_ID As Long, unit_id As Long, Width As Double, Optional ByVal mLenth As Double = 0, Optional ByVal mhight As Double = 0) As Double
   'Dim StrSQL  As String
   Dim mIsPriceIsLenthW As Boolean
   Dim mIsPriceIsLenth As Boolean
   Dim mIsPriceIsLenthWH As Boolean
   Dim StrSQL As String
   Dim mFinalPrice As Double
     Dim rs As ADODB.Recordset
     Dim mPrice As Double
    Set rs = New ADODB.Recordset
     StrSQL = "SELECT     IsNull(IsPriceIsPerview,0) IsPriceIsPerview ,IsNull(IsPriceIsLenthW,0) IsPriceIsLenthW , IsNull(IsPriceIsMatrix,0) IsPriceIsMatrix ,  IsNull(IsPriceIsLenth,0) IsPriceIsLenth ,IsNull(IsPriceIsLenthWH,0) IsPriceIsLenthWH "
StrSQL = StrSQL & " From dbo.TblItems"
StrSQL = StrSQL & "  Where (IsNull(IsPriceIsPerview,0) =1 Or IsNull(IsPriceIsLenthW,0) =1 Or IsNull(IsPriceIsLenth,0) =1   Or IsNull(IsPriceIsLenthWH,0) =1  Or IsNull(IsPriceIsMatrix,0) =1)   and ItemID = " & Item_ID
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs.EOF Then
    GetItemPriceByWitdth2 = 0
    Exit Function
End If

    
If (rs!IsPriceIsLenthW & "") Then
    mIsPriceIsLenthW = True
    mPrice = GetItemPrice(val(Item_ID), , val(unit_id))
    If Width = 0 Then Width = 1
    If mLenth = 0 Then mLenth = 1
    '«÷«›… 20 ﬂ«‰  «ÕœÏ ÿ·»«  «·„›—Ê‘«   „ «Ìﬁ«›Â« «·«‰
    'Width = (Width * (mLenth + 20))
    Width = (Width * (mLenth))
    
     mFinalPrice = mPrice * Width
    mPrice = val(mFinalPrice)
End If


If (rs!IsPriceIsLenth & "") Then
    mIsPriceIsLenth = True
    If mLenth = 0 Then mLenth = 1
    Width = (mLenth)
  '  txtQty1 = mLenth * val(txtCountItem2)
    mPrice = GetItemPrice(val(Item_ID), , val(unit_id))
    mFinalPrice = mPrice * mLenth
    mPrice = val(mFinalPrice)
End If

If (rs!IsPriceIsLenthWH & "") Then
    mIsPriceIsLenthWH = True
    If mLenth = 0 Then mLenth = 1
    If mhight = 0 Then mhight = 1
    If Width = 0 Then Width = 1
    Width = (mLenth) * Width * mhight
  '  txtQty1 = mLenth * val(txtCountItem2)
    mPrice = GetItemPrice(val(Item_ID), , val(unit_id))
    mFinalPrice = mPrice * Width
    mPrice = val(mFinalPrice)
End If

If (rs!IsPriceIsMatrix & "") Then
    mIsPriceIsLenthWH = True
        
    
    Width = GetPriceFromMatrix(Item_ID, val(unit_id), mhight, Width)
  '  txtQty1 = mLenth * val(txtCountItem2)
    mPrice = GetItemPrice(val(Item_ID), , val(unit_id))
    mFinalPrice = Width ' mPrice * Width
    mPrice = val(mFinalPrice)
    GetItemPriceByWitdth2 = mPrice
    Exit Function
End If


 StrSQL = "SELECT     dbo.Fn_GetPriceItem(" & Item_ID & ", " & Width & ") AS WidthPrice  "
StrSQL = StrSQL & " From dbo.TblItems"
StrSQL = StrSQL & "  Where (IsNull(IsPriceIsPerview,0) =1 Or IsNull(IsPriceIsLenthW,0) =1 Or IsNull(IsPriceIsLenth,0) =1 Or IsNull(IsPriceIsLenthWH,0) =1)  and ItemID = " & Item_ID
Set rs = New ADODB.Recordset
 
  
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If (rs.RecordCount) > 0 Then
       If Not mIsPriceIsLenthW Then
            GetItemPriceByWitdth2 = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width '/ 100
            ' «·Ê÷⁄ «·ﬁœÌ„ ﬂ‰« »‰ﬁ”„ ⁄·Ï 100
            'GetItemPriceByWitdth2 = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width / 100
        Else
            GetItemPriceByWitdth2 = IIf(IsNull(rs("WidthPrice").value), 0, val((rs("WidthPrice").value) & "")) * Width '/ 10000
        End If
        
        'Dim mPrice As Double
        If GetItemPriceByWitdth2 = 0 Then
            GetItemPriceByWitdth2 = mPrice
        End If
'        If val(rs!widthPrice & "") = 0 Then
'            GetItemPriceByWitdth2 = TxtPrice
'          '  txtQty1 = GetItemPriceByWitdth2
'        End If
Else
GetItemPriceByWitdth2 = 0
    End If

End Function
 
Public Function GetPriceFromMatrix(ByVal mItemNo As Long, unit_id As Long, Width As Double, Optional ByVal mLenth As Double = 0) As Double
Dim rsDummy2 As New ADODB.Recordset


    s = "Select  TblItemsUnits.UnitFactor from TblItemsUnits where ItemID = " & mItemNo & " and UnitID = " & unit_id
    rsDummy2.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy2.EOF Then
        Width = Width * val(rsDummy2!UnitFactor & "")
        mLenth = mLenth * val(rsDummy2!UnitFactor & "")
    End If
    
s = "Select * from tblItemsmatrix where ItemID = " & mItemNo & " and Value = " & Width
If Not ChkFildName("A" & mLenth, "tblItemsmatrix") Then GetPriceFromMatrix = 0: Exit Function

Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    
   
        GetPriceFromMatrix = val(rsDummy("A" & mLenth) & "")
   
Else
    GetPriceFromMatrix = 0: Exit Function
End If

End Function

Private Function ChkFildName(FieldName As String, tablename As String) As Boolean

Dim Conn As Object
Dim rs As New ADODB.Recordset
'Dim fieldName As String
'Dim tableName As String



' «” ⁄·«„ SQL ·· Õﬁﬁ „‰ ÊÃÊœ «·Õﬁ· ›Ì «·ÃœÊ·
Dim SQLQuery As String
SQLQuery = "SELECT * FROM " & tablename

' › Õ Recordset
rs.Open SQLQuery, Cn, adOpenKeyset, adLockReadOnly

' «· Õﬁﬁ „‰ ÊÃÊœ «·Õﬁ·
Dim fieldFound As Boolean
fieldFound = False
Dim fld As Object

For Each fld In rs.Fields
    If UCase(fld.Name) = UCase(FieldName) Then
        fieldFound = True
        Exit For
    End If
Next fld

' ≈€·«ﬁ «·« ’«· Ê«·—ÌﬂÊ—œ ” 
rs.Close


' ⁄—÷ ‰ ÌÃ… «· Õﬁﬁ
If fieldFound Then
    ChkFildName = True
Else
    ChkFildName = False
End If

End Function
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
If SystemOptions.PriceWithVAT = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) Then
'VatGrid.Rows = 1: If Not LblValueAdded Is Nothing Then LblValueAdded.Caption = "": Exit Sub
End If

If UCase(frmname) = "FRMSALEBILL" Then
    If chkTaxExempt.value = vbChecked Then
         VatGrid.rows = 1
         LblValueAdded.Caption = "": Exit Sub
    End If
End If

Dim mm As Double
If SystemOptions.PriceWithVAT = True And (Me.GridTrans = InvoiceTransaction) Then
If mTaxTobacco <> 0 Then
Else
VatGrid.rows = 1
End If
If Not LblValueAdded Is Nothing Then LblValueAdded.Caption = "": Exit Sub
End If

 If (Me.TxtModFlag.Text = "R" Or Me.TxtModFlag.Text = "" Or val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code"))) = 0) And posDelete = False Then Exit Sub
 Dim Percentg As Double
Dim LngItemID As Double
If StoreName Is Nothing Then Exit Sub
If True = True Then
'If TransType = 9 And ReturnSales = True Then
If ReturnTyp = 2 Then
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
Percentg = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")))
If Not txtManulaVat Is Nothing Then
            If val(txtManulaVat.Text) <> 0 Then
                    Percentg = val(txtManulaVat.Text)
            End If
End If
If Not txtManulaVat Is Nothing Then
            If val(txtManulaVat.Text) <> 0 Then
                     Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")) = val(txtManulaVat.Text)
            End If
End If

Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Round((Percentg * (val(val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) - CalculteDiscountLin(LongRow)))) / 100, 2)
If AllItems = 0 Then
FillGridValuAdded TransType, LongRow
End If
Else
LngItemID = val(Me.Grid.TextMatrix(LongRow, Grid.ColIndex("Code")))
If CheckItemWithOutVAT(LngItemID) = True Then
Percentg = -1
Else
If SystemOptions.AllItemInVAT = True Then

Percentg = val(cCompanyInfo.VATItems)
Else
Percentg = PercentgValueAdded(DtpBillDate, val(StoreName.BoundText), LngItemID, TransType)


End If

If Not txtManulaVat Is Nothing Then
            If val(txtManulaVat.Text) <> 0 Then
                    Percentg = val(txtManulaVat.Text)
            End If
End If
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
 
  If Not XPTxtDiscountVal Is Nothing Then
       If SystemOptions.DiscountByQtyOnly And val(XPTxtDiscountVal.Text) <> 0 Then
            Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".###")
       End If
    Else
        Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu")))) - CalculteDiscountLin(LongRow)))) / 100, ".###")
    End If
                            
    
 Else
    
    Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = (Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu")))) - CalculteDiscountLin(LongRow)))) / 100

   
End If

            If Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = "." Then
             Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = 0
            End If

Else

'       ''////////////////////
If SystemOptions.PriceWithVAT = True Then


'TxtValueAdded = (SngTemp + SngTemp1) * 0.05
 'SngTemp = SngTemp / (1 + (Percetage / 100))
    'xx = (val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) / (1 + (Percentg / 100)))
    
    If Not (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) Then
     
            Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".###")
    
    Else
    
    If mTaxTobacco <> 0 Then
                mm = val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) - ((val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) / 2 / 1.075) * 2)
                Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = mm
        Else
    
            Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = (val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) / (1 + (Percentg / 100))) * Percentg / 100
        End If
    End If
   ' Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percetage * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".###")
Else
    If m_GridTrans2 = 2 Then
       Dim mVVV
        mVVV = val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) - val(Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))) / (1 + Percentg / 100)
        Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format(((val(val((mVVV))))), ".###")
    Else
        Grid.TextMatrix(LongRow, Grid.ColIndex("Vat")) = Format((Percentg * (val(val((Grid.TextMatrix(LongRow, Grid.ColIndex("Valu"))))))) / 100, ".###")
    End If
End If
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
  
Public Function CheckItemWithOutVAT(Optional ItemID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "Select ItemWithOutVAT from TblItems where ItemID=" & ItemID & " and  (NOT (ItemWithOutVAT IS NULL)) AND (ItemWithOutVAT = 1 or ItemWithOutVAT=2)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckItemWithOutVAT = True
Else
CheckItemWithOutVAT = False
End If
End Function

   
  Public Function CalculteDiscountLin(Optional i As Long)
Dim DblDiscountTotal As Double
Dim TotalDiscountPerLine As Double
Dim total As Double
If LBLGross Is Nothing Then Exit Function
If LBLGross = "" Then
LBLGross = 0
End If
With Grid
     If .rows > 1 Then
            total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
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
           If TotalDiscountPerLine > val(.TextMatrix(i, .ColIndex("Valu"))) Then
                TotalDiscountPerLine = .TextMatrix(i, .ColIndex("Valu")) / (total) * DblDiscountTotal
           End If
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
.rows = Grid.rows
For i = 1 To Grid.rows - 1
If Grid.ColIndex("Vat") <> -1 Then 'salim1503
If val(Grid.TextMatrix(i, Grid.ColIndex("Vat"))) <> 0 And val(Grid.TextMatrix(i, Grid.ColIndex("Code"))) <> 0 Then
End If
.TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("Code")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Code"))
.TextMatrix(i, .ColIndex("Name")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Name"))
.TextMatrix(i, .ColIndex("select")) = True
.TextMatrix(i, .ColIndex("Vat")) = Grid.TextMatrix(i, Grid.ColIndex("Vat"))
.TextMatrix(i, .ColIndex("Vatyo")) = Grid.TextMatrix(i, Grid.ColIndex("Vatyo"))
If Not txtManulaVat Is Nothing Then
            If val(txtManulaVat.Text) <> 0 Then
                    .TextMatrix(i, .ColIndex("Vatyo")) = val(txtManulaVat.Text)
                     Grid.TextMatrix(i, Grid.ColIndex("Vatyo")) = val(txtManulaVat.Text)
            End If
End If

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
Dim Start As Integer
cunt = 0
If VatGrid Is Nothing Then Exit Sub
With VatGrid
.rows = Grid.rows
 Ros = LongRow
Start = LongRow
If mTaxTobacco <> 0 Then Start = 1
For i = Start To Ros
If Grid.ColIndex("Vat") <> -1 Then 'salim1503
cunt = Start
'RetriveItems val(Grid.TextMatrix(i, Grid.ColIndex("Code"))), Name, code
.TextMatrix(cunt, .ColIndex("index")) = cunt
.TextMatrix(cunt, .ColIndex("Code")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Code"))
.TextMatrix(cunt, .ColIndex("Name")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Name"))
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
            If .Row + 1 < .rows Then
                .Row = Row + 1
            ElseIf .Row + 1 = .rows Then
                .rows = .rows + 1
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

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
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
    Dim StrSQL  As String
    Dim mTotal As Double
    Dim rsDummy As New ADODB.Recordset
    With Me.Grid
        
        If val(Grid.TextMatrix(Row, Grid.ColIndex("DiscountType"))) = 4 And .ColKey(Col) <> "Select" And GridTrans <> ReturnSalling Then
            Cancel = True
            Exit Sub
        End If

        Select Case .ColKey(Col)
        
            Case "Valu", "StillPumbQty"
                Cancel = True
                Exit Sub
            Case "RequestTypeName"
                If val(Grid.TextMatrix(Row, Grid.ColIndex("ItemBalance"))) < val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) Then

                    Cancel = True
                Else
                
                End If
            Case "CashQty", "MadaQty", "VisaQty"
            '.TextMatrix(row, .ColIndex("Cash")) = val(.TextMatrix(row, .ColIndex("CashQty"))) * val(.TextMatrix(row, .ColIndex("Price")))
            mTotal = val(.TextMatrix(Row, .ColIndex("CashQty"))) + val(.TextMatrix(Row, .ColIndex("MadaQty"))) + val(.TextMatrix(Row, .ColIndex("VisaQty"))) + val(.TextMatrix(Row, .ColIndex("DeferredQty")))
            If mTotal > val(.TextMatrix(Row, .ColIndex("Count"))) Then
                MsgBox "Total quantities do not match the actual quantity."
                '.TextMatrix(row, Col) = mTotal - val(.TextMatrix(row, Col))
                .TextMatrix(Row, Col) = 0
                .TextMatrix(Row, Col + 1) = 0
                Grid_AfterEdit Row, Col
                
            End If
            Case "ShipedQty"
                Cancel = True

            Case "ISSUEDQTY"
                Cancel = True
            Case "OrderLineId"
                Cancel = True
            Case "OldID"
                Cancel = True
            Case "TotalInvoiceQty"
                Cancel = True

            Case "RecivedShippedQty"
                Cancel = True
            Case "TotallRecivedShippedQty"
                Cancel = True

            Case "AddSerials"

                If .cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexUnchecked Then
                    Cancel = True
                End If

            Case "Name"

                If .ColComboList(.ColIndex("Name")) = "" Then
                    Cancel = True
                End If

            Case "HaveSerial"
                Cancel = True

            Case "Count"

                If .cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexChecked Then
                    Cancel = True
                Else
                    Cancel = False
                End If

            Case "Serial", "guaranteeTime"

                If .cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexChecked Then
                    Cancel = False
                Else
                    Cancel = True
                End If
            Case "DiscountType", "DiscountVal"
                If SystemOptions.NotEditDiscountLine Then
                    Cancel = True
                End If
            Case "Price"
                Dim mPrice As Double
                
                
                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
        
                    StrSQL = "Select Emp_ID,CustomerId,FromDate,Todate  from TblCustomerContract "
                    StrSQL = StrSQL & " Left Outer join TblCustomerContractDetails On TblCustomerContractDetails.TblCustomerContractD = TblCustomerContract.TblCustomerContractD"
                    StrSQL = StrSQL & " Where customerid = " & val(Customer.BoundText)
                    StrSQL = StrSQL & " and TblCustomerContractDetails.ItemID =" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                    StrSQL = StrSQL & " and " & SQLDate(DtpBillDate.value, True) & "   BETWEEN FromDate and Todate"
        
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummy.EOF Then
                        Cancel = True
                    Else
                        Grid.EditMaxLength = 10
                    End If

                End If
    
                If Me.GridTrans = InvoiceTransaction Then
                           
                    Dim rsDummyPrice As New ADODB.Recordset
                                    
                    StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                                    
                    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                    StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId and DefaultUnit = 1"
                    StrSQL = StrSQL + " Where TblItems.ItemID=" & val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))
                    rsDummyPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummyPrice.EOF Then
                        If val(rsDummyPrice!MinSelingPrice & "") > 0 And val(Grid.TextMatrix(Row, Grid.ColIndex("Price"))) < val(rsDummyPrice!MinSelingPrice & "") Then
                            MsgBox "·« Ì„ﬂ‰ «·»Ì⁄ »«ﬁ· „‰ «·Õœ «·«œ‰Ï ··”⁄—"
                            Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(rsDummyPrice!MinSelingPrice & "")
                            Exit Sub
                        ElseIf val(rsDummyPrice!MaxSelingPrice & "") > 0 And val(Grid.TextMatrix(Row, Grid.ColIndex("Price"))) > val(rsDummyPrice!MaxSelingPrice & "") Then
                            MsgBox "·« Ì„ﬂ‰ «·»Ì⁄ »«⁄·Ï  „‰ «·Õœ «·«⁄·Ï ··”⁄—"
                                            
                            Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(rsDummyPrice!MaxSelingPrice & "")
                                            
                            Exit Sub
                        End If
                                    
                    End If

                End If
    
                If Me.GridTrans = ReturnSalling Then
                    If SystemOptions.NotEditSalesRetPrice Then
                        Cancel = True
                    End If
                End If
    
                If Me.GridTrans = internalorder Then
                    If SystemOptions.NotEditInternalPrice Then
                        Cancel = True
                    End If
                End If
                mPrice = Grid.TextMatrix(Row, Grid.ColIndex("Price"))
                If Me.GridTrans = purchaseorder Or Me.GridTrans = purchaseorderrequest Or Me.GridTrans = PurchaseTransaction Then
                    Exit Sub
                End If

                If SystemOptions.CanChangePriceUpOnly Then
                 Set rsDummyPrice = New ADODB.Recordset
                                    
                    StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID,TBLiTEMSUnits.UnitSalesPrice, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                                    
                    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                    StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId "
                    'and DefaultUnit = 1"
                    StrSQL = StrSQL + " Where TblItems.ItemID=" & val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))
                    StrSQL = StrSQL + "  and TBLiTEMSUnits.UnitID=" & val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                    rsDummyPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummyPrice.EOF Then
                        If mPrice < val(rsDummyPrice!UnitSalesPrice & "") Then
                             Grid.TextMatrix(Row, Grid.ColIndex("Price")) = val(rsDummyPrice!UnitSalesPrice & "")
                        End If
                    End If
                    
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

                            If SystemOptions.NotEditDiscountLine Then
                                Cancel = True
                            Else
                                Cancel = False
                            End If
                        Else
                            Cancel = True
                        End If
                    Else
                        Cancel = False
                    End If
                Else
                    Cancel = False
                End If
                If .ColKey(Col) = "DiscountType" Or .ColKey(Col) = "DiscountVal" Then

                    If SystemOptions.NotEditDiscountLine Then
                        Cancel = True
                    Else
                        Cancel = False

                    End If

                End If
                Select Case .ColKey(Col)
                    Case "StoreID2"

                        If .TextMatrix(Row, .ColIndex("MixNo")) <> "" Then
                            Cancel = True
                        Else
                            Cancel = False
                        End If
                    Case "Price"
                        If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
        
                            StrSQL = "Select Emp_ID,CustomerId,FromDate,Todate  from TblCustomerContract "
        
                            StrSQL = StrSQL & " Left Outer join TblCustomerContractDetails On TblCustomerContractDetails.TblCustomerContractD = TblCustomerContract.TblCustomerContractD"
                            StrSQL = StrSQL & " Where customerid = " & val(Customer.BoundText)
                            StrSQL = StrSQL & " and TblCustomerContractDetails.ItemID =" & val(Grid.TextMatrix(Row, Grid.ColIndex("Code")))
                            StrSQL = StrSQL & " and " & SQLDate(DtpBillDate.value, True) & "   BETWEEN FromDate and Todate"
                            Set rsDummy = New ADODB.Recordset
                            rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                            If Not rsDummy.EOF Then
                                Cancel = True
                            Else
                                Grid.EditMaxLength = 10
                            End If

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
    Dim total As Double
    Dim RealTotal As Double
    Dim TotalNet As Double
    Dim DblDiscountTotal As Double
    Dim DblRowTotal As Double
        Dim DblRowTotal2 As Double
      Dim DblRowTotal3 As Double
      
      Dim DblProfitTotal As Double
      Dim mProfitValue  As Double
      Dim mNetValue As Double
      Dim mTaxTobaccoTotal As Double
          If SystemOptions.PriceWithVAT = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) Then VatGrid.rows = 1: If Not LblValueAdded Is Nothing Then LblValueAdded.Caption = ""
        Dim DblRowTotals As Double
        
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim DblQty As Double
    Dim StrSQL  As String
    Dim i As Long
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
        RowNum = .FixedRows
        
        
        Dim mRwosTotal As Long
        
        mRwosTotal = .rows - 1
        
        Dim mFixedRow As Long
        mFixedRow = 1
        If BolOneRow Then
            If LngRow <> 0 Then RowNum = LngRow: mFixedRow = LngRow: mRwosTotal = LngRow Else LngRow = 0
        Else
        
        End If
        For RowNum = mFixedRow To mRwosTotal
            
           ' .TextMatrix(LngRow, .ColIndex("UnitPrice")) = .TextMatrix(LngRow, .ColIndex("Price"))
'            If SystemOptions.IsGeometricProportions = True And val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) <> 0 Then
'                If val(.TextMatrix(LngRow, .ColIndex("Count"))) <> 0 Then
'                    .TextMatrix(LngRow, .ColIndex("Price")) = val(.TextMatrix(LngRow, .ColIndex("UnitPrice"))) / val(.TextMatrix(LngRow, .ColIndex("Count")))
'                End If
'            End If

            If .ColIndex("UnitPrice") <> -1 Then
                        If SystemOptions.IsGeometricProportions = True And val(.TextMatrix(RowNum, .ColIndex("UnitPrice"))) <> 0 Then
                        If val(.TextMatrix(RowNum, .ColIndex("Count"))) <> 0 Then
                            .TextMatrix(RowNum, .ColIndex("Price")) = val(.TextMatrix(RowNum, .ColIndex("UnitPrice"))) * val(.TextMatrix(RowNum, .ColIndex("NoCount"))) / val(.TextMatrix(RowNum, .ColIndex("Count")))
                            End If
                        End If
            End If
            
                If m_GridTrans2 = 2 Then
                    
                    '.TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa"))) + val(.TextMatrix(RowNum, .ColIndex("Deferred"))), SystemOptions.SysDefCurrencyForamt)
                    'Wael
                    .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa"))), SystemOptions.SysDefCurrencyForamt)
                 Else
                    .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                 End If

            
                    
                    
                    
                        
            If UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
                .TextMatrix(LngRow, .ColIndex("AmountHComm")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("AmountH"))) / 100, SystemOptions.SysDefCurrencyForamt)
            End If
                    
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
                 
                 If m_GridTrans2 = 2 Then
                    
                    'wael
                    '.TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa"))) + val(.TextMatrix(RowNum, .ColIndex("Deferred"))), SystemOptions.SysDefCurrencyForamt)
                    .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa"))), SystemOptions.SysDefCurrencyForamt)
                 Else
                    .TextMatrix(RowNum, .ColIndex("Valu")) = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                 End If
            '    End If
                
                

End If



            LngCurItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
            If Grid.ColIndex("UnitID") <> -1 Then 'salim1503

            LngUnitID = val(.cell(flexcpData, RowNum, .ColIndex("UnitID")))    'val(.TextMatrix(RowNum, .ColIndex("UnitID")))
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
'Wael   ”Ì „ «Ìﬁ«› „Ê÷Ê⁄ «·÷—Ì»… ·ÕÌ‰ «·‰Ÿ— «·ÌÂ«
        If .rows > 1 Then
'            If m_GridTrans2 = 2 Then
'                total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu")) / 1.15
'            Else
'                total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
'            End If
'            total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
            
            
            
             total = 0
            For i = 1 To .rows - 1
                total = total + (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("Price"))))
            Next
                                        
            RealTotal = total
        End If
    If m_GridTrans2 = 2 Then
        frmsalebill6.CalcValueGrids
        total = val(frmsalebill6.mMada) + val(frmsalebill6.mVisa) + val(frmsalebill6.mCash)
        RealTotal = total
    End If
        If Not Me.LblTotalAll Is Nothing Then
            If mTaxTobacco <> 0 Then
                mTaxTobaccoTotal = total / 2 / 1.075
                total = mTaxTobaccoTotal
                Me.LblTotalAll.Caption = Round(total, SystemOptions.SysDefCurrencyForamt)
            Else
                Me.LblTotalAll.Caption = Round(total, SystemOptions.SysDefCurrencyForamt)
            End If
            '          Me.LblTotalAll.Caption = Val(Format(total, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
        End If
    
        DblDiscountTotal = 0
  DblCommissionTotal = 0
  
        Dim discountvalue As Double
        Dim DblPrice As Double
        Dim mValProfit As Double
                                     
                             Dim Quantity As Double
                            Dim Price As Double
                            Dim TotalValue As Double
                            Dim TotalQty As Double
                            'Dim discountvalue As Double
                            Dim TotalShahnPerLine As Double

        If .ColIndex("DiscountType") <> -1 Then

For RowNum = mFixedRow To mRwosTotal
            'For RowNum = 1 To .Rows - 1
                
                If m_GridTrans2 = 2 Then
                    'Wael
                    'DblRowTotal = val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa"))) + val(.TextMatrix(RowNum, .ColIndex("Deferred")))
                    DblRowTotal = val(.TextMatrix(RowNum, .ColIndex("Cash"))) + val(.TextMatrix(RowNum, .ColIndex("Mada"))) + val(.TextMatrix(RowNum, .ColIndex("Visa")))
                Else
                
                    DblRowTotal = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                End If
                
                DblPrice = val(.TextMatrix(RowNum, .ColIndex("Price")))
          
          
          
           If Not XPTxtDiscountVal Is Nothing Then
                    If SystemOptions.DiscountByQtyOnly And val(XPTxtDiscountVal.Text) <> 0 Then
                        If Not CboDiscount_Type Is Nothing Then
                            .TextMatrix(RowNum, .ColIndex("DiscountType")) = CboDiscount_Type.ListIndex + 1
                        Else
                            .TextMatrix(RowNum, .ColIndex("DiscountType")) = 2
                        End If
                    End If
            End If
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
                             
                             
                            ' «·Õ’Ê· ⁄·Ï «·ﬁÌ„
                            
                            If Not XPTxtDiscountVal Is Nothing Then
                            If SystemOptions.DiscountByQtyOnly And val(XPTxtDiscountVal.Text) <> 0 Then
                                       TotalShahnPerLine = 0
                                       
                                        Quantity = val(.TextMatrix(RowNum, .ColIndex("Count")))
                                        Price = DblPrice
                                        TotalValue = val(LblTotalAll.Caption)
                                        TotalQty = val(LblTotalQty.Caption)
                                        discountvalue = val(XPTxtDiscountVal.Text)
                                        
                                        '  ÿ»Ìﬁ «·‘—ÿ
                                        If SystemOptions.DiscountByQtyOnly = False Then
                                            '  Ê“Ì⁄ «·Œ’„ »‰«¡ ⁄·Ï «·ﬂ„Ì… ›ﬁÿ
                                            If TotalQty <> 0 Then
                                                If Not CboDiscount_Type Is Nothing Then
                                                    If CboDiscount_Type.ListIndex = 1 Then
                                                        TotalShahnPerLine = Round(((Quantity / TotalQty) * discountvalue / 100), SystemOptions.SysDefCurrencyForamt + 1)
                                                    Else
                                                        TotalShahnPerLine = Round(((Quantity / TotalQty) * discountvalue), SystemOptions.SysDefCurrencyForamt + 1)
                                                    End If
                                                Else
                                                    TotalShahnPerLine = Round(((Quantity / TotalQty) * discountvalue), SystemOptions.SysDefCurrencyForamt + 1)
                                                End If
                                            End If
                                        Else
                                            If TotalValue <> 0 Then
                                            '  Ê“Ì⁄ «·Œ’„ »‰«¡ ⁄·Ï «·ﬂ„Ì… * «·”⁄—
                                                If Not CboDiscount_Type Is Nothing Then
                                                    If CboDiscount_Type.ListIndex = 2 Then
                                                        TotalShahnPerLine = Round((((Price * Quantity) / TotalValue) * discountvalue), SystemOptions.SysDefCurrencyForamt + 1)
                                                    Else
                                                        TotalShahnPerLine = Round((((Price * Quantity) / TotalValue) * discountvalue), SystemOptions.SysDefCurrencyForamt + 1)
                                                    End If
                                                
                                                Else
                                                    TotalShahnPerLine = Round((((Price * Quantity) / TotalValue) * discountvalue), SystemOptions.SysDefCurrencyForamt + 1)
                                                End If
                                                discountvalue = Round(TotalShahnPerLine, SystemOptions.SysDefCurrencyForamt + 1)
                                            End If
                                        End If
                                

                            End If
                            End If
                                        If SystemOptions.DiscountByQtyOnly = True And TotalShahnPerLine <> 0 Then
                                            .TextMatrix(RowNum, .ColIndex("DiscountVal")) = TotalShahnPerLine
                                            If Not CboDiscount_Type Is Nothing Then
                                                .TextMatrix(RowNum, .ColIndex("DiscountType")) = CboDiscount_Type.ListIndex + 1
                                                If CboDiscount_Type.ListIndex + 1 = 3 Then
                                                    GoTo DiscountLine
                                                End If
                                            Else
                                                .TextMatrix(RowNum, .ColIndex("DiscountType")) = 2
                                            End If
                                         Else
                                            DblDiscountTotal = DblDiscountTotal + val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                                            discountvalue = val(.TextMatrix(RowNum, .ColIndex("DiscountVal")))
                                        End If
            
                                Case 3
DiscountLine:
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

 .TextMatrix(RowNum, .ColIndex("Valu")) = Round(DblRowTotal - discountvalue, SystemOptions.SysDefCurrencyForamt)
'**********************************************************************
        If .ColIndex("ProfitType") <> -1 Then
            DblRowTotal = Round(val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))), SystemOptions.SysDefCurrencyForamt)
                '.TextMatrix(RowNum, .ColIndex("ProfitValue")) = "" 'DblRowTotal
                .TextMatrix(RowNum, .ColIndex("NetProfit")) = ""

            Select Case val(.TextMatrix(RowNum, .ColIndex("ProfitType")))

            Case 1
                .TextMatrix(RowNum, .ColIndex("ProfitValue")) = "" 'DblRowTotal
                .TextMatrix(RowNum, .ColIndex("NetProfit")) = ""
                

            Case 2
                If val(.TextMatrix(RowNum, .ColIndex("ProfitValue"))) <> 0 Then
                .TextMatrix(RowNum, .ColIndex("NetProfit")) = DblRowTotal + val(.TextMatrix(RowNum, .ColIndex("ProfitValue")))
                 DblProfitTotal = DblProfitTotal + val(.TextMatrix(RowNum, .ColIndex("NetProfit")))
                End If

            

            Case 3
            
               
                


'                                    .TextMatrix(RowNum, .ColIndex("Valu")) = (DblRowTotal) * (1 - (val(.TextMatrix(RowNum, .ColIndex("DiscountVal"))) / 100))
                mProfitValue = val(.TextMatrix(RowNum, .ColIndex("ProfitValue")))
                DblProfitTotal = DblProfitTotal + ((val(.TextMatrix(RowNum, .ColIndex("ProfitValue"))) * val(DblRowTotal)) / 100)
                mValProfit = ((val(.TextMatrix(RowNum, .ColIndex("ProfitValue"))) * DblRowTotal) / 100)
               ' DblProfitTotal = DblProfitTotal + val(.TextMatrix(RowNum, .ColIndex("NetProfit")))
               .TextMatrix(RowNum, .ColIndex("NetProfit")) = mValProfit + DblRowTotal
        End Select

            '.TextMatrix(RowNum, .ColIndex("NetProfit")) = DblProfitTotal + mValProfit
        End If
        
        
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
If .rows > 1 Then
'Wael   ”Ì „ «Ìﬁ«› „Ê÷Ê⁄ «·÷—Ì»… ·ÕÌ‰ «·‰Ÿ— «·ÌÂ«
        If m_GridTrans2 = 2 Then
            total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu")) / 1.15
        Else
            total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
        End If
        total = .Aggregate(flexSTSum, 1, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
        
        
       
          If m_GridTrans2 = 2 Then
            frmsalebill6.CalcValueGrids
            total = val(frmsalebill6.mMada) + val(frmsalebill6.mVisa) + val(frmsalebill6.mCash)
            
       
        End If
End If
                        If Me.GridTrans = 3 Or Me.GridTrans <> ShowPrice And Me.GridTrans <> Template And Me.GridTrans <> Destruction And Me.GridTrans <> Returntransaction And Me.GridTrans <> StockSettlement Then
            
                            'Õ”«» ﬁÌ„… «·Œ’„
                                                 If Not CboDiscount_Type Is Nothing Then
                                                                    If CboDiscount_Type.ListIndex = 1 Or CboDiscount_Type.ListIndex = 3 Then
                                                                    
                                                                        DblDiscountTotal = DblDiscountTotal + val(TxtDiscount_Val.Text)
                                                                     
                                                                        total = RealTotal - DblDiscountTotal
                                                                    Else
                                                
                                                                                    If CboDiscount_Type.ListIndex = 2 Then
                                                                                                If TxtDiscount_Val.Text <> "" Then
                                                                                                    DblDiscountTotal = DblDiscountTotal + ((val(TxtDiscount_Val.Text) * total) / 100)
                                                                                                
                                                                                                    '  total = RealTotal * Val(1 - (TxtDiscount_Val.text / 100))
                                                                                                    total = RealTotal - DblDiscountTotal
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
        TotalNet = total

        '÷—»Ì… Œ’„ Ê≈÷«›…
        If Not Me.TxtAddTax Is Nothing Then
            If val(Me.TxtAddTax.Text) > 0 Then
                Dim DblTemp As Double

                If Not Me.LblTaxAddValue Is Nothing Then
                    LblTaxAddValue.Caption = ((val(Me.TxtAddTax.Text) * total) / 100)
                End If

                TotalNet = total * (1 - val((Me.TxtAddTax.Text)) / 100)
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
                        LblTaxSalesValue.Caption = ((val(Me.TxtTaxValue.Text) * total) / 100)
                    End If

                    TotalNet = TotalNet + ((val(Me.TxtTaxValue.Text) * total) / 100)
                End If
            End If
        End If

        '÷—»Ì… Œœ„…
        If Not Me.TxtServiceTax Is Nothing Then
            If val(Me.TxtServiceTax.Text) <> 0 Then
                If Not Me.LblTaxServiceValue Is Nothing Then
                    Me.LblTaxServiceValue.Caption = ((val(Me.TxtServiceTax.Text) * total) / 100)
                End If

                TotalNet = TotalNet + ((val(Me.TxtServiceTax.Text) * total) / 100)
            End If
        End If

        TotalNet = Round(TotalNet, SystemOptions.SysDefCurrencyForamt)
        txtTotal.Text = TotalNet
    Else
        txtTotal.Text = Round(total, SystemOptions.SysDefCurrencyForamt)
    End If

    If Not BolSavae = False Then
        If TxtModFlag.Text = "N" Or TxtModFlag.Text = "E" Then
            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) > val(total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
                Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Calculate = False
                Exit Function
            End If

            If val(TxtValueCash.Text) + val(TxtValueDelay.Text) + val(TxtValuechque.Text) < val(total) Then
                Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
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
            .TextMatrix(Crow, .ColIndex("Count")) = IIf(SystemOptions.DefaultQtyTrans = 0, "", SystemOptions.DefaultQtyTrans) '1
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
                            For i = 1 To .rows - 1
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
Dim StoreID As Integer
CaseNo = 0

With Grid
StoreID = val(.TextMatrix(1, .ColIndex("StoreID2")))
For i = 1 To .rows - 1
        If val(.TextMatrix(i, .ColIndex("Code"))) <> 0 And (val(.TextMatrix(i, .ColIndex("StoreID2")))) <> StoreID Then
        StoreID = val(.TextMatrix(i, .ColIndex("StoreID2")))
                        If CaseNo < 2 Then
                        CaseNo = CaseNo + 1
                        End If
        End If
Next i
If StoreID <> val(StoreName.BoundText) And CaseNo = 0 Then
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

Dim intDefStore As Integer
GetUserData user_id, , , intDefStore


Set cCompanyInfo = New ClsCompanyInfo
    m_ShowQtyToolTip = True
If storename1 <> -1 Or ItemByCus = 1 Then GoTo xxx

    With Grid



'≈÷«›… «·„÷Œ« 
    If m_GridTrans2 = 2 Then
  
        If Me.Grid.ColIndex("PumpName") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "PumpName"
            .ColPosition(.ColIndex("PumpName")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("PumpName")) = "«·„÷Œ…"
            Else
                .TextMatrix(0, .ColIndex("PumpName")) = "Pump Name"
            End If
        Else
            .ColAlignment(.ColIndex("PumpName")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("PumpName")) = flexAlignRightCenter
            StrSQL = "Select * From tblPumpType Order by Name"
            Set RsNote = New ADODB.Recordset
            RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrList = .BuildComboList(RsNote, "Name", "Id")
    
            If StrList <> "" Then
                .ColComboList(.ColIndex("PumpName")) = "|" & StrList
            End If
        
        End If
        
        
        
        
        
        
        If Me.Grid.ColIndex("PrevQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "PrevQty"
            .ColPosition(.ColIndex("PrevQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("PrevQty")) = "ﬁ—«¡… ”«»ﬁ…"
            Else
                .TextMatrix(0, .ColIndex("PrevQty")) = "PrevQty"
            End If
    
            .ColAlignment(.ColIndex("PrevQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("PrevQty")) = flexAlignRightCenter
        End If
        ''''''
          If Me.Grid.ColIndex("CurrentQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "CurrentQty"
            .ColPosition(.ColIndex("CurrentQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("CurrentQty")) = "ﬁ—«¡… Õ«·Ì…"
            Else
                .TextMatrix(0, .ColIndex("CurrentQty")) = "CurrentQty"
            End If
    
            .ColAlignment(.ColIndex("CurrentQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("CurrentQty")) = flexAlignRightCenter
        End If
        ''''''
        
        If Me.Grid.ColIndex("StillPumbQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "StillPumbQty"
            .ColPosition(.ColIndex("StillPumbQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("StillPumbQty")) = "Still PumbQty"
            Else
                .TextMatrix(0, .ColIndex("StillPumbQty")) = "StillPumbQty"
            End If
                
            .ColAlignment(.ColIndex("StillPumbQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("StillPumbQty")) = flexAlignRightCenter
            
        End If
        
        
         If Me.Grid.ColIndex("CashQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "CashQty"
            .ColPosition(.ColIndex("CashQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("CashQty")) = "Cash Qty"
            Else
                .TextMatrix(0, .ColIndex("CashQty")) = "Cash Qty"
            End If
    
            .ColAlignment(.ColIndex("CashQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("CashQty")) = flexAlignRightCenter
            
        End If
        
        
         If Me.Grid.ColIndex("Cash") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Cash"
            .ColPosition(.ColIndex("Cash")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Cash")) = "ﬂ«‘"
            Else
                .TextMatrix(0, .ColIndex("Cash")) = "Cash"
            End If
    
            .ColAlignment(.ColIndex("Cash")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Cash")) = flexAlignRightCenter
            
        End If
        
        
                
        If Me.Grid.ColIndex("MadaQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "MadaQty"
            .ColPosition(.ColIndex("MadaQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("MadaQty")) = "Mada Qty"
            Else
                .TextMatrix(0, .ColIndex("MadaQty")) = "Mada Qty"
            End If
    
            .ColAlignment(.ColIndex("MadaQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("MadaQty")) = flexAlignRightCenter
        End If
        

        
        If Me.Grid.ColIndex("Mada") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Mada"
            .ColPosition(.ColIndex("Mada")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Mada")) = "„œÏ"
            Else
                .TextMatrix(0, .ColIndex("Mada")) = "Mada"
            End If
    
            .ColAlignment(.ColIndex("Mada")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Mada")) = flexAlignRightCenter
        End If
        
                                   
        If Me.Grid.ColIndex("VisaQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "VisaQty"
            .ColPosition(.ColIndex("VisaQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("VisaQty")) = "VisaQty"
            Else
                .TextMatrix(0, .ColIndex("VisaQty")) = "VisaQty"
            End If
    
            .ColAlignment(.ColIndex("VisaQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("VisaQty")) = flexAlignRightCenter
        End If
        
                                 
        If Me.Grid.ColIndex("Visa") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Visa"
            .ColPosition(.ColIndex("Visa")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Visa")) = "Visa"
            Else
                .TextMatrix(0, .ColIndex("Visa")) = "Visa"
            End If
    
            .ColAlignment(.ColIndex("Visa")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Visa")) = flexAlignRightCenter
        End If
        
        
        If Me.Grid.ColIndex("DeferredQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "DeferredQty"
            .ColPosition(.ColIndex("DeferredQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("DeferredQty")) = "DeferredQty"
            Else
                .TextMatrix(0, .ColIndex("DeferredQty")) = "DeferredQty"
            End If
    
            .ColAlignment(.ColIndex("DeferredQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("DeferredQty")) = flexAlignRightCenter
        End If
        
        
        If Me.Grid.ColIndex("Deferred") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Deferred"
            .ColPosition(.ColIndex("Deferred")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Deferred")) = "¬Ã·"
            Else
                .TextMatrix(0, .ColIndex("Deferred")) = "¬Ã·"
            End If
    
            .ColAlignment(.ColIndex("Deferred")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Deferred")) = flexAlignRightCenter
        End If


        If Me.Grid.ColIndex("AmountH") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "AmountH"
            .ColPosition(.ColIndex("AmountH")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("AmountH")) = "«·Â··« "
            Else
                .TextMatrix(0, .ColIndex("AmountH")) = "«·Â··« "
            End If
    
            .ColAlignment(.ColIndex("AmountH")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("AmountH")) = flexAlignRightCenter
        End If
        
        If Me.Grid.ColIndex("AmountHComm") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "AmountHComm"
            .ColPosition(.ColIndex("AmountHComm")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("AmountHComm")) = "«·›—Êﬁ« "
            Else
                .TextMatrix(0, .ColIndex("AmountHComm")) = "«·›—Êﬁ« "
            End If
    
            .ColAlignment(.ColIndex("AmountHComm")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("AmountHComm")) = flexAlignRightCenter
        End If
        
        If Me.Grid.ColIndex("IsOther") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "IsOther"
            .ColPosition(.ColIndex("IsOther")) = .Cols - 1
            .ColDataType(.ColIndex("IsOther")) = flexDTBoolean
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("IsOther")) = "··€Ì—"
            Else
                .TextMatrix(0, .ColIndex("IsOther")) = "··€Ì—"
            End If
    
            .ColAlignment(.ColIndex("IsOther")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("IsOther")) = flexAlignRightCenter
        End If
        
          
        If Me.Grid.ColIndex("Account_CodeComm") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Account_CodeComm"
            .ColPosition(.ColIndex("Account_CodeComm")) = .Cols - 1
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Account_CodeComm")) = ""
            Else
                .TextMatrix(0, .ColIndex("Account_CodeComm")) = "··€Ì—"
            End If
    
            .ColAlignment(.ColIndex("Account_CodeComm")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Account_CodeComm")) = flexAlignRightCenter
        End If
        
         If Me.Grid.ColIndex("Account_Code") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "Account_Code"
            .ColPosition(.ColIndex("Account_Code")) = .Cols - 1
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Account_Code")) = ""
            Else
                .TextMatrix(0, .ColIndex("Account_Code")) = "··€Ì—"
            End If
    
            .ColAlignment(.ColIndex("Account_Code")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("Account_Code")) = flexAlignRightCenter
        End If
        
        If Me.Grid.ColIndex("DetailsPump") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "DetailsPump"
            
            .ColPosition(.ColIndex("DetailsPump")) = .Cols - 1
            .ColComboList(.ColIndex("DetailsPump")) = "|..."
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("DetailsPump")) = ""
            Else
                .TextMatrix(0, .ColIndex("DetailsPump")) = "··€Ì—"
            End If
    
            .ColAlignment(.ColIndex("DetailsPump")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("DetailsPump")) = flexAlignRightCenter
        End If
        
  
  
 


    ElseIf m_GridTrans2 = 1 Then
        .ColHidden(.ColIndex("PumpName")) = True
    
    End If
        
       
        
        
            
        If SystemOptions.UserInterface = EnglishInterface Then
            .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignLeftCenter
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


        If .ColIndex("ProfitType") <> -1 Then
                '.TextMatrix(0, .ColIndex("ProfitType")) = "ProfitType"
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignRightCenter
                    .ColComboList(.ColIndex("ProfitType")) = "#1;·«ÌÊÃœ |#2;—»ÕÌ… »ﬁÌ„…|#3;—»ÕÌ… »‰”»…|"
                ElseIf SystemOptions.UserInterface = EnglishInterface Then
                    .ColComboList(.ColIndex("ProfitType")) = "#1;NO |#2;Value|#3;Percent profit|#"
                End If

            End If
        If SystemOptions.UserInterface = ArabicInterface Then
            .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignRightCenter
            .ColComboList(.ColIndex("ItemCase")) = "#1;ÃœÌœ|#2;„” ⁄„·"
        Else
            .ColComboList(.ColIndex("ItemCase")) = "#1;New|#2;Used"
        End If

     If .ColIndex("ItemCodeSearch") <> -1 Then
                .TextMatrix(0, .ColIndex("ItemCodeSearch")) = "Quick Search"
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
         If .ColIndex("RequestTypeName") <> -1 Then
                .TextMatrix(0, .ColIndex("RequestTypeName")) = "‰Ê⁄ «·ÿ·»"
                .ColComboList(.ColIndex("RequestTypeName")) = "#1;ÿ·» ‘—«¡|#2;ÿ·»  ÕÊÌ· „Œ“‰Ì"
            Else
        
            End If

            If .ColIndex("ItemBalance2") <> -1 Then
                .TextMatrix(0, .ColIndex("ItemBalance2")) = "«·—’Ìœ ›Ï «·„Œ“‰ «·„ÿ·Ê· «· ÕÊÌ· „‰Â"
                
            Else
        
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



       ' .Cols = .Cols + 1
       ' .ColKey(.Cols - 1) = "TotalQtyIn"
       ' .ColPosition(.ColIndex("TotalQtyIn")) = .Cols - 1
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .TextMatrix(0, .ColIndex("TotalQtyIn")) = "TotalQtyIn"
'        Else
'            .TextMatrix(0, .ColIndex("TotalQtyIn")) = "TotalQtyIn"
'        End If
'
'        .ColAlignment(.ColIndex("TotalQtyIn")) = flexAlignRightCenter
'        .FixedAlignment(.ColIndex("TotalQtyIn")) = flexAlignRightCenter



'        .Cols = .Cols + 1
'        .ColKey(.Cols - 1) = "TotalQtyOut"
'        .ColPosition(.ColIndex("TotalQtyOut")) = .Cols - 1
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .TextMatrix(0, .ColIndex("TotalQtyOut")) = "TotalQtyOut"
'        Else
'            .TextMatrix(0, .ColIndex("TotalQtyOut")) = "TotalQtyOut"
'        End If
'
'        .ColAlignment(.ColIndex("TotalQtyOut")) = flexAlignRightCenter
'        .FixedAlignment(.ColIndex("TotalQtyOut")) = flexAlignRightCenter
'


      '  .Cols = .Cols + 1
      '  .ColKey(.Cols - 1) = "TotalBalance"
      '  .ColPosition(.ColIndex("TotalBalance")) = .Cols - 1
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .TextMatrix(0, .ColIndex("TotalBalance")) = "TotalBalance"
'        Else
'            .TextMatrix(0, .ColIndex("TotalBalance")) = "TotalBalance"
'        End If
'
'        .ColAlignment(.ColIndex("TotalBalance")) = flexAlignRightCenter
'        .FixedAlignment(.ColIndex("TotalBalance")) = flexAlignRightCenter




      '  .Cols = .Cols + 1
      '  .ColKey(.Cols - 1) = "TotalBalance"
      '  .ColPosition(.ColIndex("TotalBalance")) = .Cols - 1
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .TextMatrix(0, .ColIndex("TotalBalance")) = "TotalBalance"
'        Else
'            .TextMatrix(0, .ColIndex("TotalBalance")) = "TotalBalance"
'        End If
'
'        .ColAlignment(.ColIndex("TotalBalance")) = flexAlignRightCenter
'        .FixedAlignment(.ColIndex("TotalBalance")) = flexAlignRightCenter
'
'
'
'
'        .Cols = .Cols + 1
'        .ColKey(.Cols - 1) = "TotalCostBalance"
'        .ColPosition(.ColIndex("TotalCostBalance")) = .Cols - 1
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .TextMatrix(0, .ColIndex("TotalCostBalance")) = "TotalBalance"
'        Else
'            .TextMatrix(0, .ColIndex("TotalCostBalance")) = "TotalBalance"
'        End If
'
'        .ColAlignment(.ColIndex("TotalCostBalance")) = flexAlignRightCenter
'        .FixedAlignment(.ColIndex("TotalCostBalance")) = flexAlignRightCenter
'
'


        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "NLC"
        .ColPosition(.ColIndex("NLC")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("NLC")) = "NLC"
        Else
            .TextMatrix(0, .ColIndex("NLC")) = "NLC"
        End If

        .ColAlignment(.ColIndex("NLC")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("NLC")) = flexAlignRightCenter




        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "TotalCost"
        .ColPosition(.ColIndex("TotalCost")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("TotalCost")) = "TotalCost"
        Else
            .TextMatrix(0, .ColIndex("TotalCost")) = "TotalCost"
        End If

        .ColAlignment(.ColIndex("TotalCost")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("TotalCost")) = flexAlignRightCenter




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
        
        
        
         .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "chkTaxExempt"
        .ColPosition(.ColIndex("chkTaxExempt")) = .Cols - 1
        .ColDataType(.Cols - 1) = flexDTBoolean
        .ColHidden(.Cols - 1) = True
    If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("chkTaxExempt")) = "„⁄›Ï „‰ «·÷—Ì»…"
        Else
            .TextMatrix(0, .ColIndex("chkTaxExempt")) = "Tax Exempt"
        End If
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
          
          
                .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "OldID"
        .ColPosition(.ColIndex("OldID")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("OldID")) = "OldID"
        Else
            .TextMatrix(0, .ColIndex("OldID")) = "OldID"
        End If

        .ColAlignment(.ColIndex("OldID")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("OldID")) = flexAlignRightCenter
        
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
            .TextMatrix(0, .ColIndex("ItemQty")) = "ﬂ„Ì… 2"
        Else
            .TextMatrix(0, .ColIndex("ItemQty")) = "qty2"
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
                                                       
  .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "thickness"
        .ColPosition(.ColIndex("thickness")) = .Cols - 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("thickness")) = "«·”„ﬂ"
        Else
            .TextMatrix(0, .ColIndex("thickness")) = "thickness"
        End If
        

        .ColAlignment(.ColIndex("thickness")) = flexAlignRightCenter
        .FixedAlignment(.ColIndex("thickness")) = flexAlignRightCenter
                                                             
                                                       
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
    StrSQL = StrSQL & "     where 1 = 1 "
     If SystemOptions.usertype <> UserAdmin Then
        StrSQL = StrSQL & "     and   TblEmployee.BranchId = " & Current_branch
     End If
     If SystemOptions.IsMashghal Then
            StrSQL = StrSQL & "     and       IsNull(TblEmployee.chkShowTasks,0) = 1"
     End If
    'StrSQL = StrSQL & "     where  ( JobTypeName like '%⁄«„·%'  or JobTypeNamee like '%Labor%' )  "
 


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
    

StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemWithOutVAT, Groups.chkTaxExempt, dbo.TblItems.GroupID, dbo.TblItems.HaveSerial, dbo.TblItems.LastUpdate, "
StrSQL = StrSQL & "                       dbo.TblItems.UserID, dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblItems.RequestLimit, dbo.TblItems.CustomerPrice, dbo.TblItems.DealerPrice,"
StrSQL = StrSQL & "                       dbo.TblItems.HaveGuarantee, dbo.TblItems.GuaranteeValue, dbo.TblItems.GuaranteeType, dbo.TblItems.IsArchive, dbo.TblItems.ItemType, dbo.TblItems.AssbliedItem,"
StrSQL = StrSQL & "                       dbo.TblItems.RelatedItem, dbo.TblItems.ItemComment, dbo.TblItems.ItemCase, dbo.TblItems.ItemMaking, dbo.TblItems.ItemMakingNew, dbo.TblItems.code,"
StrSQL = StrSQL & "                       dbo.TblItems.Branch_NO, dbo.TblItems.Fullcode, dbo.TblItems.prifix, dbo.TblItems.PartNo, dbo.TblItems.CostPrice, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                       dbo.TblItems.DefaultSupplier, dbo.TblItems.itemSerials, dbo.TblItems.BinLocation, dbo.TblItems.minvalueqty, dbo.TblItems.MaxValueqty, dbo.TblItems.FreeQty,"
StrSQL = StrSQL & "                       dbo.TblItems.barCodeNO, dbo.TblItems.CatlogNO, dbo.TblItems.FactoryNO, dbo.TblItems.TemplateID, dbo.TblItems.ItemMaxDiscount, dbo.TblItems.OverHead,"
StrSQL = StrSQL & "                       dbo.TblItems.Wight , dbo.TblItems.Content, dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.maxRecivePeriod"

StrSQL = "Select Groups.chkTaxExempt,TblItems.* "
StrSQL = StrSQL & "  FROM         dbo.TblItems Inner join Groups On TblItems.GroupId = Groups.GroupId  where IsArchive=0      "
  If Me.GridTrans = InvoiceTransaction And SystemOptions.ShowOnlyItemsOfSales = True Then
    StrSQL = StrSQL & "  and dbo.TblItems.GroupID in( Select  Groups.GroupId  "
    StrSQL = StrSQL & " From dbo.Groups  WHERE  (ISNULL(POSGroup, 0) = 1))"
  End If
    If UCase(frmname) = "FRMSALEBILL3" Then
        StrSQL = StrSQL & " and (groups.BranchID =  " & my_branch & " ) "
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
                 
                            If val(StoreName.BoundText) = 0 Then
                                StrSQL = GETLINKSQL(intDefStore, 0, Me.TxtModFlag)
                            Else
                                StrSQL = GETLINKSQL(val(StoreName.BoundText), 0, Me.TxtModFlag)
                             End If
                 
                                 GoTo order_By
                             
                                 End If
                 
                 End If
        
        
        
        
End If



If SystemOptions.WorkWithLINKEDiActivity = True Then

            '     If Me.TxtModFlag = "N" Or Me.TxtModFlag = "E" Then
                
                 
            '                     If 1 = 1 Then
                     
                 '29 12 2020
                             StrSQL = GETLINKSQLByActivity(0, 0, "˙˙˙XX")
                 
                                 GoTo order_By
                             
            '                     End If
                 
            '     End If
        
        
        
        
End If

'**********************************


    If Me.GridTrans = Destruction Then
        StrSQL = StrSQL
        GoTo order_By
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ShowPrice Then
        StrSQL = StrSQL + " and (TblItems.AssbliedItem=1 OR TblItems.AssbliedItem=0)"
    ElseIf Me.GridTrans = ProductionOrder Then
      '  StrSQL = StrSQL + " and (TblItems.ItemMakingNew=1)"
    Else
      '  StrSQL = StrSQL + " and (TblItems.AssbliedItem=0)"
    End If

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
        StrSQL = StrSQL + " AND (TblItems.ItemType=0 OR TblItems.ItemType=1)"
    Else
      '  StrSQL = StrSQL + " AND (TblItems.ItemType=0)"
    End If
'StrSQL = StrSQL + " AND (ItemID<=3500)"

order_By:

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order By ItemName"
    Else
        StrSQL = StrSQL + " Order By ItemNamee"
    End If
    If UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
        StrSQL = " Select Groups.chkTaxExempt,TblItems.*   FROM         dbo.TblItems Inner join Groups On TblItems.GroupId = Groups.GroupId  WHERE itemid IN (SELECT"
        StrSQL = StrSQL & " tblPumpType.ItemID"
        StrSQL = StrSQL & " FROM tblPumpType)"
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

    With Grid
        If Me.Grid.ColIndex("UsedQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "UsedQty"
            .ColPosition(.ColIndex("UsedQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("UsedQty")) = "«·ﬂ„Ì… «·„” Œœ„…"
            Else
                .TextMatrix(0, .ColIndex("UsedQty")) = "UsedQty"
            End If
    
            .ColAlignment(.ColIndex("UsedQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("UsedQty")) = flexAlignRightCenter
        End If
        If Me.Grid.ColIndex("StillQty") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "StillQty"
            .ColPosition(.ColIndex("StillQty")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("StillQty")) = "«·ﬂ„Ì… «·„ »ﬁÌ…"
            Else
                .TextMatrix(0, .ColIndex("StillQty")) = "StillQty"
            End If
    
            .ColAlignment(.ColIndex("StillQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("StillQty")) = flexAlignRightCenter
            
        End If
       
        If Me.Grid.ColIndex("OrderLineId") = -1 Then
            .Cols = .Cols + 1
            .ColKey(.Cols - 1) = "OrderLineId"
            .ColPosition(.ColIndex("OrderLineId")) = .Cols - 1
    
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("OrderLineId")) = "—ﬁ„ «·”ÿ— ›Ï «·Õ—ﬂ…"
            Else
                .TextMatrix(0, .ColIndex("OrderLineId")) = "OrderLineId"
            End If
    
            .ColAlignment(.ColIndex("StillQty")) = flexAlignRightCenter
            .FixedAlignment(.ColIndex("StillQty")) = flexAlignRightCenter
            
        End If
       
  
    End With
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
If UCase(frmname) = "FRMSALEBILL6" Then
    ShowHidCol
End If
  
If storename1 = -1 Then
    GetMySetting
End If
    Exit Sub
ErrTrap:
    'Stop
    'Resume
End Sub

Public Sub ShowHidCol()
Dim i As Long

With Grid
For i = 1 To .Cols - 1
    
    If m_GridTrans2 = 1 Then
        .ColHidden(.ColIndex("PumpName")) = True
    Else
        .ColHidden(i) = True
    End If
Next
    If m_GridTrans2 = 2 Then
        .ColHidden(.ColIndex("PumpName")) = False
        .ColHidden(.ColIndex("Code")) = False
        .ColHidden(.ColIndex("Name")) = False
        .ColHidden(.ColIndex("PrintName")) = False
        .ColHidden(.ColIndex("Count")) = False
        .ColHidden(.ColIndex("Price")) = False
        .ColHidden(.ColIndex("Valu")) = False
        .ColHidden(.ColIndex("PrevQty")) = False
        .ColHidden(.ColIndex("CurrentQty")) = False
        .ColHidden(.ColIndex("Cash")) = False
        .ColHidden(.ColIndex("Visa")) = False
        .ColHidden(.ColIndex("Mada")) = False
        .ColHidden(.ColIndex("Deferred")) = False
        .ColHidden(.ColIndex("StillPumbQty")) = False
        .ColFormat(.ColIndex("Price")) = "#,###.##########"
        .ColFormat(.ColIndex("Valu")) = "#,###.##########"
        
        .ColHidden(.ColIndex("CashQty")) = False
        .ColHidden(.ColIndex("VisaQty")) = False
        .ColHidden(.ColIndex("MadaQty")) = False
        .ColHidden(.ColIndex("DeferredQty")) = False
        
        .ColHidden(.ColIndex("StoreId2")) = False
        
        .ColHidden(.ColIndex("CusID2")) = False
        
        .ColHidden(.ColIndex("AmountH")) = False
        .ColHidden(.ColIndex("AmountHComm")) = False
        .ColHidden(.ColIndex("DetailsPump")) = False
        .ColHidden(.ColIndex("Account_CodeComm")) = False
        .ColHidden(.ColIndex("Account_Code")) = False
        .ColHidden(.ColIndex("IsOther")) = False
        
        
                .ColHidden(.ColIndex("PumpName")) = False
        .ColHidden(.ColIndex("Code")) = False
        .ColHidden(.ColIndex("Name")) = False
        
        .ColHidden(.ColIndex("Count")) = False
        .ColHidden(.ColIndex("Price")) = False
        .ColHidden(.ColIndex("Valu")) = False
        .ColHidden(.ColIndex("PrevQty")) = False
        .ColHidden(.ColIndex("CurrentQty")) = False
        .ColHidden(.ColIndex("StoreId2")) = False
        
        .ColHidden(.ColIndex("PumpName")) = False
        .ColHidden(.ColIndex("Code")) = False
        .ColHidden(.ColIndex("Name")) = False
        
        .ColHidden(.ColIndex("Count")) = False
        .ColHidden(.ColIndex("Price")) = False
        .ColHidden(.ColIndex("Valu")) = False
        .ColHidden(.ColIndex("PrevQty")) = False
        .ColHidden(.ColIndex("CurrentQty")) = False
        .ColHidden(.ColIndex("Cash")) = False
        .ColHidden(.ColIndex("Visa")) = False
        .ColHidden(.ColIndex("Mada")) = False
        .ColHidden(.ColIndex("Deferred")) = False
        .ColHidden(.ColIndex("StoreId2")) = False
        
        .ColHidden(.ColIndex("CusID2")) = False
                .ColHidden(.ColIndex("AmountH")) = True
        .ColHidden(.ColIndex("AmountHComm")) = False
        .ColHidden(.ColIndex("DetailsPump")) = False
        .ColHidden(.ColIndex("Account_CodeComm")) = True
        .ColHidden(.ColIndex("Account_Code")) = True
        .ColHidden(.ColIndex("IsOther")) = False
        


        
        
    Else
       
        
                
        .ColHidden(.ColIndex("AmountH")) = True
        .ColHidden(.ColIndex("AmountHComm")) = True
        .ColHidden(.ColIndex("DetailsPump")) = True
        .ColHidden(.ColIndex("Account_CodeComm")) = True
        .ColHidden(.ColIndex("Account_Code")) = True
        .ColHidden(.ColIndex("IsOther")) = True
                .ColFormat(.ColIndex("Price")) = "#,###.##"
        .ColFormat(.ColIndex("Valu")) = "#,###.##"
.ColHidden(.ColIndex("PrintName")) = False
    End If
     .ColHidden(.ColIndex("Code")) = False
        .ColHidden(.ColIndex("Name")) = False
        .ColHidden(.ColIndex("PrintName")) = False
        .ColHidden(.ColIndex("Count")) = False
        .ColHidden(.ColIndex("Price")) = False
        .ColHidden(.ColIndex("Valu")) = False
End With

End Sub

 
Private Sub Grid_BeforeSort(ByVal Col As Long, _
                            Order As Integer)
    Dim i As Integer

    With Grid

        If .TextMatrix(.rows - 1, .ColIndex("Code")) = "" Then
            .RemoveItem .rows - 1
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
                Case "DetailsPump"
                   
                    frmsalebill6.LngRow = Row

                    frmsalebill6.LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                    FrmItemShowDet.mIndex = 1
                    'FrmItemShowDet.mItemID = val(Grid.cell(flexcpData, row, Grid.ColIndex("Name"))) 'LngItemID
                    
                    FrmItemShowDet.mItemId = IIf((Grid.TextMatrix(Row, Grid.ColIndex("Code")) = ""), 0, val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))))
                    FrmItemShowDet.mUnitId = val(Grid.cell(flexcpData, Row, Grid.ColIndex("UnitID")))
                    FrmItemShowDet.strDisplay = Trim(.TextMatrix(Row, .ColIndex("DetailsPump")))
                    FrmItemShowDet.DisplayGrid
                    FrmItemShowDet.txtStillQty = Trim(.TextMatrix(Row, .ColIndex("StillPumbQty")))
                    Load FrmItemShowDet
                    
                    
                    FrmItemShowDet.show vbModal
            Case "AddSerials"

                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Then
                    If .cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexUnchecked Then
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
    StrSQL = "Select * from tblItems Where 1 = 1 "
    'If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = ShowPrice Then
    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or Me.GridTrans = internalorder Or Me.GridTrans = ShowPrice Or Me.GridTrans = ReturnSalling Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
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
  If UCase(frmname) = "FRMSALEBILL3" Then
        StrItem = StrItem & " and groupId In (Select Groups.GroupId from groups where groups.BranchID =  " & my_branch & " ) "
  End If
   ' StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
'        If StrItem = "" Then
'            StrItem = StrSQL
'        End If
    If UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
         StrSQL = " Select ItemId,ItemCode from tblItems where itemid In (Select tblPumpType.ItemId from tblPumpType)"
 
        fill_combo Me.DCboItemCode, StrSQL
        
                 StrSQL = " Select ItemId,ItemName from tblItems where itemid In (Select tblPumpType.ItemId from tblPumpType)"
 
        fill_combo Me.DCboItemName, StrSQL
    Else
        Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, , storename1, , , StrItem
        Dcombos.GetItemsNames Me.DCboItemName, -1, -1, , , storename1, , , StrItem
    End If
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

    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = internalorder Or Me.GridTrans = ShowPrice Or Me.GridTrans = salespricelist Or Me.GridTrans = salespricelistRequest Then
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

 
Private Sub Grid_DblClick()
    Dim itemcode As String
    Dim ItemName As String

    With Me.Grid

        If .Row = -1 Then
            Exit Sub
        End If
        If .Col = -1 Then
            Exit Sub
        End If
        
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
              If SystemOptions.WorkWithItemsDetails = False Then GoTo nome
       ItemDetailsNewIgia .Row, .Col
       Exit Sub
       ''
       If SystemOptions.WorkWithItemsDetails = False Then GoTo nome
            FrmItemsDetails.show
            FrmItemsDetails.TxtItemID = .TextMatrix(.Row, .ColIndex("Code"))
            GetItemData val(.TextMatrix(.Row, .ColIndex("Code"))), itemcode, ItemName
            FrmItemsDetails.TxtItemCode = itemcode
            FrmItemsDetails.txtItemName = ItemName
          
            FrmItemsDetails.Txtqty = .TextMatrix(.Row, .ColIndex("Count"))
            FrmItemsDetails.TxtUnitID = .cell(flexcpData, .Row, .ColIndex("UnitID"))
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
nome:
        End If

    End With

    If Me.TxtModFlag.Text = "R" Then

        With Me.Grid

            If .Row = -1 Then
                Exit Sub
            End If
            If .Col = -1 Then
                Exit Sub
            End If
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
                    FrmSelectData.TxtItemCode.Text = val(.TextMatrix(.Row, .ColIndex("Name")))

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
                    .rows = .rows + 1
                    .Row = .Row + 1
                End If
         
                .Col = .ColIndex("Code")
          
        
      End If
      
GoTo ll
End If

        
            If .Col = .ColIndex("Price") And val(.TextMatrix(.Row, .ColIndex("Price"))) <> 0 Then
                currentrow = .Row
                '.Rows = .Rows + 1
                If .Row + 1 <= .rows - 1 Then
                .Row = .Row + 1
                End If
                If currentrow = .Row Then
                    .rows = .rows + 1
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
           If .ColKey(.Col) = "Account_Name" Or .ColKey(.Col) = "Account_Serial2" Then
        
                Account_search.show
                Account_search.case_id = 808080
                Exit Sub
            End If
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

            If LngMouseRow = -1 Then
                Exit Sub
            End If
            If .Col = -1 Then
                Exit Sub
            End If
            If val(.TextMatrix(LngMouseRow, .ColIndex("Name"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("Name")))

                If Me.Grid.cell(flexcpChecked, LngMouseRow, .ColIndex("HaveSerial")) = flexChecked Then
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

                If LngRow = -1 Then
                    Exit Sub
                End If
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
    Dim Rs1       As ADODB.Recordset
    
    Dim StrSQL    As String
    Dim MyStrList As String
    Dim LngItemID As Long
    Dim mTotal As Double
    With Me.Grid
        '    If Me.GridTrans = InvoiceTransaction Then

        Select Case .ColKey(Col)
          Case "CashQty", "MadaQty", "VisaQty"
            '.TextMatrix(row, .ColIndex("Cash")) = val(.TextMatrix(row, .ColIndex("CashQty"))) * val(.TextMatrix(row, .ColIndex("Price")))
           
            If mTotal > val(.TextMatrix(Row, .ColIndex("Count"))) Then
               ' MsgBox "Total quantities do not match the actual quantity."
                
            End If
            
'          Case "MadaQty"
'          ', "MadaQty", "VisaQty", "DeferredQty"
'            .TextMatrix(row, .ColIndex("Mada")) = val(.TextMatrix(row, .ColIndex("MadaQty"))) * val(.TextMatrix(row, .ColIndex("Price")))
'
'          Case "VisaQty"
'          ', "MadaQty", "VisaQty", "DeferredQty"
'            .TextMatrix(row, .ColIndex("Visa")) = val(.TextMatrix(row, .ColIndex("VisaQty"))) * val(.TextMatrix(row, .ColIndex("Price")))
            
          Case "Cash", "Mada", "Visa", "Deferred", "DeferredQty"
                
                
                Cancel = True
'                If mTotal > val(.TextMatrix(row, .ColIndex("Valu"))) Then
'                    MsgBox "·« Ì„ﬂ‰  ŒÿÏ ﬁÌ„… «·„÷Œ…"
'                    .TextMatrix(row, Col) = 0
'                    .TextMatrix(row, Col) = 0
'                    DoEvents
'                    .TextMatrix(row, Col) = 0
'                End If
            Case "Count"
     
                If SystemOptions.WorkWithItemsDetails = True And (Me.GridTrans <> NewGard) Then
                    Cancel = True
                End If
                If UCase(frmname) = "FRMSALEBILL5" Then
                    If GridTrans <> ReturnSalling Then
                        Cancel = True
                    End If
                End If
     
            Case "RequestTypeName"
                If val(Grid.TextMatrix(Row, Grid.ColIndex("ItemBalance2"))) < val(Grid.TextMatrix(Row, Grid.ColIndex("Count"))) Then
                    Cancel = True
                    '.ColComboList(.ColIndex("RequestTypeName")) = "#1;ÿ·» ‘—«¡|"
                    '      .ColComboList(.ColIndex("RequestTypeName")) = "#1;ÿ·» ‘—«¡|#2;ÿ·»  ÕÊÌ· „Œ“‰Ì"
                Else
                    '     .ColComboList(.ColIndex("RequestTypeName")) = "#1;ÿ·» ‘—«¡|#2;ÿ·»  ÕÊÌ· „Œ“‰Ì"
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
            
            Case "Account_Name", "Account_Serial2"
              '  If .ColComboList(Col) = "" Then
'                    StrSQL = "SELECT '1' as [" & Format(Now, "yyyymmhhsshhnn") & "],   ACCOUNTS.Account_Code,  REPLACE(REPLACE(REPLACE(ACCOUNTS.Account_Name, CHAR(10), ''), CHAR(13), ''), CHAR(9), '')  As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
'
'                    If SystemOptions.SysDataBaseType = AccessDataBase Then
'                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
'                    Else
'                        StrSQL = StrSQL + " And(   ACCOUNTS.Block=0 or  ACCOUNTS.Block is null)    and ACCOUNTS.last_account=1"
'                    End If
'
'                    StrSQL = StrSQL & GetAccountByBarnchUser
'                    StrSQL = StrSQL & GetAccountCodeHiding
'                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
'
'                    Set rsAccount = New ADODB.Recordset
'                    rsAccount.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly
'                    StrComboList = .BuildComboList(rsAccount, "RootName,ParentName,*FirstName", "Account_Code")      ' "RootName,ParentName,*FirstName", "Account_Code")
'                    If StrComboList <> "" Then
'                        StrComboList = "|" & StrComboList
'                    End If
'                    rsAccount.Close
'
'                    .ColComboList(.ColIndex("Account_Name")) = StrComboList
'                    .ColComboList(.ColIndex("Account_Serial2")) = StrComboList
'
'
'
'
'                     Dim StrList As String
'                StrSQL = "Select spht as sph , spht  as SPH From SPHTable   order by id"
'
'                '    StrSQL = "Select * From TblItemsColors Order by ColorName"
'                Dim RsNote As New ADODB.Recordset
'                RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'                StrList = .BuildComboList(RsNote, "SPH", "SPH")
'
'                If StrList <> "" Then
'                    .ColComboList(.ColIndex("LSPH")) = "|" & StrList
'                    .ColComboList(.ColIndex("RSPH")) = "|" & StrList
'                End If
'
'                RsNote.Close
'                End If
'
              
              
              
              
                    StrSQL = "SELECT ACCOUNTS.Account_Code,  ACCOUNTS.Account_Name FirstName,Account_Serial"
                    StrSQL = StrSQL & " FROM ACCOUNTS Where ACCOUNTS.Account_Code <>'r' "
                    
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(   ACCOUNTS.Block=0 or  ACCOUNTS.Block is null)    and ACCOUNTS.last_account=1"
                    End If
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"

                    Set rsAccount = New ADODB.Recordset
                    rsAccount.Open StrSQL, Cn, adOpenForwardOnly, adLockReadOnly
                    If .ColKey(Col) = "Account_Serial2" Then
                        StrComboList = .BuildComboList(rsAccount, "Account_Serial", "Account_Code")    ' "RootName,ParentName,*FirstName", "Account_Code")
                    Else
                        StrComboList = .BuildComboList(rsAccount, "FirstName", "Account_Code")    ' "RootName,ParentName,*FirstName", "Account_Code")
                    End If
                    If StrComboList <> "" Then
                        StrComboList = "|" & StrComboList
                    End If
                    rsAccount.Close
                
                    .ColComboList(.ColIndex("Account_Name")) = StrComboList
                    .ColComboList(.ColIndex("Account_Serial2")) = StrComboList
              ' End If

            Case "LSPH", "RSPH"

                '*********************************
                '    .ColAlignment(.ColIndex("LSPH")) = flexAlignRightCenter
                '    .FixedAlignment(.ColIndex("LSPH")) = flexAlignRightCenter
    
                Dim StrList As String
                StrSQL = "Select spht as sph , spht  as SPH From SPHTable   order by id"
            
                '    StrSQL = "Select * From TblItemsColors Order by ColorName"
                Dim RsNote As New ADODB.Recordset
                RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrList = .BuildComboList(RsNote, "SPH", "SPH")

                If StrList <> "" Then
                    .ColComboList(.ColIndex("LSPH")) = "|" & StrList
                    .ColComboList(.ColIndex("RSPH")) = "|" & StrList
                End If

                RsNote.Close
        
                '*********************************
             
            Case "StoreIDAviName"

                '*********************************
                '    .ColAlignment(.ColIndex("LSPH")) = flexAlignRightCenter
                '    .FixedAlignment(.ColIndex("LSPH")) = flexAlignRightCenter
                If Grid.TextMatrix(Row, Grid.ColIndex("RequestTypeNo")) = "1" Then
                    sComboStore = "Select StoreId,StoreName from TblStore "
                
                Else
                    GetBalnceOfItems val(Store_ID), val(Grid.TextMatrix(Row, Grid.ColIndex("Code"))), Row
                End If
                StrSQL = sComboStore
            
                '    StrSQL = "Select * From TblItemsColors Order by ColorName"
                '  Dim RsNote As New ADODB.Recordset
        
                RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
                StrList = .BuildComboList(RsNote, "StoreName", "StoreId")

                If StrList <> "" Then
                    .ColComboList(.ColIndex("StoreIDAviName")) = "|" & StrList
           
                End If

                RsNote.Close
       
            Case "LCYL", "RCYL"

                '*********************************
                StrSQL = "Select   CLYT as CLY ,CLYT  as CLY From CLYTable   order by id"
    
                Set RsNote = New ADODB.Recordset
                RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrList = .BuildComboList(RsNote, "CLY", "CLY")

                If StrList <> "" Then
                    .ColComboList(.ColIndex("LCYL")) = "|" & StrList
                    .ColComboList(.ColIndex("RCYL")) = "|" & StrList
                End If

                RsNote.Close
        
                '*********************************
            Case "RAXIS", "LAXIS"

                '*********************************
                StrSQL = "Select aix,aix From aixTable   order by id"
    
                Set RsNote = New ADODB.Recordset
                RsNote.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrList = .BuildComboList(RsNote, "aix", "aix")

                If StrList <> "" Then
                    .ColComboList(.ColIndex("RAXIS")) = "|" & StrList
                    .ColComboList(.ColIndex("LAXIS")) = "|" & StrList
                End If

                RsNote.Close
        
                '*********************************
       
            Case "MaxUnitID"
                ' Cancel = True
  
                '       Grid.TextMatrix(Row, Grid.ColIndex("MaxQty")) = 0
                LngItemID = val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))

                If LngItemID = 0 Then
                    Cancel = True
                Else
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName,TblUnites.UnitNamee "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    
                    Set Rs1 = New ADODB.Recordset
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (Rs1.BOF Or Rs1.EOF) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MyStrList = .BuildComboList(Rs1, "UnitName", "UnitID")
                        Else
                            MyStrList = .BuildComboList(Rs1, "UnitNamee", "UnitID")
                        End If

                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("MaxUnitID")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
       
            Case "UnitID"
                LngItemID = val(Me.Grid.TextMatrix(Row, .ColIndex("Name")))

                If LngItemID = 0 Then
                    Cancel = True
                Else
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName,TblUnites.UnitNamee "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    
                    Set Rs1 = New ADODB.Recordset
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (Rs1.BOF Or Rs1.EOF) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MyStrList = .BuildComboList(Rs1, "UnitName", "UnitID")
                        Else
                            MyStrList = .BuildComboList(Rs1, "UnitNamee", "UnitID")
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
                    
                    Set Rs1 = New ADODB.Recordset
                    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (Rs1.BOF Or Rs1.EOF) Then
                   
                        MyStrList = .BuildComboList(Rs1, "CatlogName", "ID")

                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("catalog")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            Case "Price"
                If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                    Dim rsDummy As New ADODB.Recordset
                    StrSQL = "Select Emp_ID,CustomerId,FromDate,Todate  from TblCustomerContract Where CustomerId = " & val(Customer.BoundText)
        
                    StrSQL = StrSQL & " and " & SQLDate(DtpBillDate.value, True) & "   BETWEEN FromDate and Todate"
        
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummy.EOF Then
                        Cancel = True
                    Else
             
                    End If
                End If
         
            Case "project"
                StrSQL = "  SELECT     id, Project_name, Project_nameE"
                StrSQL = StrSQL & "     From dbo.Projects"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = StrSQL & " Where (Not (Project_name Is Null)) and Project_name<>N'""'"
                Else
                    StrSQL = StrSQL & " Where (Not (Project_nameE Is Null))and Project_nameE <>N'""'"
                End If
                Set Rs1 = New ADODB.Recordset
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    MyStrList = Me.Grid.BuildComboList(Rs1, "Project_name", "id")
                Else
                    MyStrList = Me.Grid.BuildComboList(Rs1, "Project_nameE", "id")
                End If

                If MyStrList <> "" Then
                    MyStrList = "|" & MyStrList
                End If

                .ComboList = MyStrList
                Rs1.Close
                Set Rs1 = Nothing
                ''//////////////
            Case "pand"
                StrSQL = " SELECT     oprid, des"
                StrSQL = StrSQL & "          From dbo.projects_des"
                StrSQL = StrSQL & "          Where (project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & " and project_id<>0)"
        
                Set Rs1 = New ADODB.Recordset
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                MyStrList = .BuildComboList(Rs1, "des", "oprid")
                If MyStrList <> "" Then
                    MyStrList = "|" & MyStrList
                End If

                .ComboList = MyStrList
                Rs1.Close
                Set Rs1 = Nothing

            Case "opera"
                StrSQL = "       SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
                StrSQL = StrSQL & "       FROM         dbo.terms_operations LEFT OUTER JOIN"
                StrSQL = StrSQL & "       dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
                StrSQL = StrSQL & "   Where (dbo.terms_operations.project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & ") And (dbo.terms_operations.ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("PandID"))) & ")"
                Set Rs1 = New ADODB.Recordset
                Rs1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                    MyStrList = .BuildComboList(Rs1, "ProcessName", "OPRIDD")
                Else
                    MyStrList = .BuildComboList(Rs1, "ProcessNameE", "OPRIDD")
                End If
                If MyStrList <> "" Then
                    MyStrList = "|" & MyStrList
                End If

                .ComboList = MyStrList
                Rs1.Close
                Set Rs1 = Nothing
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
                For i = 1 To .rows - 1
                    .TextMatrix(i, .ColIndex("ExpiryDate")) = ""
                    If val(.TextMatrix(i, .ColIndex("Code"))) <> 0 Then
                        GetItemBalanceInStore i, 1, 1, 1, , val(.TextMatrix(i, .ColIndex("Code"))), DtpBillDate.value
                        
                            
                        
                        
                        If Grid.ColIndex("RequestTypeName") <> -1 Then
                          
                            '    Grid.TextMatrix(i, Grid.ColIndex("ItemBalance")) = 0
                                GetBalnceOfItems val(Me.StoreName.BoundText), val(Grid.TextMatrix(i, Grid.ColIndex("Code"))), i, True
                         
                        End If

                    End If
                Next i
                
            End With
                                  
                       If SystemOptions.WorkWithLINKEDiTEMS = True Then
                    
                          
                              If StoreName.BoundText = "" Then
                                Exit Sub
                              End If
                                   With Grid
                                   .rows = 2
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
 

Private Sub STORENAME_Click(Area As Integer)
    'On Error GoTo ErrTrap
 Exit Sub
     Dim Dcombos As ClsDataCombos
 
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
         If SystemOptions.WorkWithLINKEDiTEMS = True Then
       '   Dcombos.GetItemsCodes Me.DCboItemCode, -1, -1, , val(storename.BoundText)
       ' Dcombos.GetItemsNames Me.DcboItemName, -1, -1, , val(storename.BoundText)
       If StoreName.BoundText = "" Then
        Exit Sub
       End If
                 
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

If mTaxTobacco <> 0 Then
    For i = 1 To Grid.rows - 1
        
            SmValu = SmValu + val(Grid.TextMatrix(i, Grid.ColIndex("Vat")))

    Next i
Else
With VatGrid
For i = 1 To .rows - 1
 If mTaxTobacco <> 0 Then
.TextMatrix(i, .ColIndex("select")) = True
End If
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then

SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
End If
Next i
End With
End If
'If SystemOptions.PriceWithVAT = True Then
'    TxtValueAdded.text = ""
'Else
         TxtValueAdded.Text = Format(SmValu, ".##")
'End If

If frmname = "frmsalebill2" And SystemOptions.PriceWithVAT = True Then
    'LblTotalAll.Caption = val(LblTotalAll) - val(TxtValueAdded.text)
End If
'LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub
Private Sub TxtAddTax_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAddTax.Text, 0)
End Sub



Private Sub TxtDiscount_Val_Change()
If frmname <> "Projects" Then
    SentTypeVAT
End If
End Sub
Private Sub GetQtyCodeFormBarCode()

    Dim mFirstPart As String
    Dim mMiddlePart As String
    Dim mLastPart As String
    Dim mResult As String
    Dim mRes As Double
    Dim DateFromBarCode As String
    Dim SupplierFromBarCode As String
    
  Dim ExpiryDateFromBarCode As Variant
    Dim mFromNo As Double
     
    '#1234101182456
    CodeFromBarCode = TxtItemCodeB1
    QtyFromBarCode = 0
    If mId(TxtItemCodeB1, 1, 1) = "#" Then
    TxtItemCodeB1 = mId(TxtItemCodeB1, 2, Len(TxtItemCodeB1) - 1)
    CodeFromBarCode = mId(TxtItemCodeB1, 1, 4)
     DateFromBarCode = mId(TxtItemCodeB1, 5, 6)
     ExpiryDateFromBarCode = mId(DateFromBarCode, 1, 2) & "/" & mId(DateFromBarCode, 3, 2) & "/20" & mId(DateFromBarCode, 5, 2)
    SupplierFromBarCode = mId(TxtItemCodeB1, 7, 3)
    End If
    
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

Private Sub TxtItemCodeB1_GotFocus()
 Me.TxtItemCodeB1.Text = ""
End Sub


Private Sub txtItemCodeSearch_GotFocus()
 Me.txtItemCodeSearch.Text = ""
End Sub


'Private Sub TxtDiscount_Val_LostFocus()
'If val(CboDiscount_Type.ListIndex) > 0 Then
'SentTypeVAT
'End If
'End Sub


Public Sub txtItemCodeSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error Resume Next
If KeyCode = vbKeyReturn Then
        Dim s As String
        Dim rsDummy As New ADODB.Recordset
        s = "Select * from tblItems Where code = N'" & Trim(txtItemCodeSearch) & "'"
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            
          TxtItemCodeB1 = rsDummy!barCodeNO & ""
          If UCase(frmname) = "FRMSALEBILL" Or UCase(frmname) = "FRMBILLBUY" Or UCase(frmname) = "FRMPO10" Then
                    DCboItemCode.BoundText = val(rsDummy!ItemID & "")
                    DCboItemName.BoundText = val(rsDummy!ItemID & "")
                    DCboItemName.SetFocus
                    DcboItemName_KeyDown KeyCode, Shift
                    
                    
            Exit Sub
          End If
            TxtItemCodeB1_KeyDown KeyCode, Shift
        End If
   
End If

End Sub
Public Sub txtItemCodeSearch2_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error Resume Next
If KeyCode = vbKeyReturn Then
        Dim s As String
        Dim rsDummy As New ADODB.Recordset
        s = "Select * from tblItems Where Fullcode = N'" & Trim(txtItemCodeSearch2) & "'"
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then

                    DCboItemCode.BoundText = val(rsDummy!ItemID & "")
                    DCboItemName.BoundText = val(rsDummy!ItemID & "")
                    DCboItemName.SetFocus
                   ' DcboItemName_KeyDown KeyCode, Shift
                    
                    
            Exit Sub
            Else
             If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«ÌÊÃœ ’‰› »Â–« «·ﬂÊœ"
            '  MsgBox "·«ÌÊÃœ ’‰› »Â–« «·ﬂÊœ", vbCritical

             Else
               'MsgBox "There is no item in this code", vbCritical
             Msg = "There is no item in this code"
             End If
             MsgBox Msg, vbCritical
    ' MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
            txtItemCodeSearch2.Text = ""
            txtItemCodeSearch2.SetFocus
          End If
   
   
End If

End Sub

Public Sub TxtItemCodeB1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyReturn Then
        '   On Error GoTo ErrTrap
        Dim StrSQL           As String
        Dim RsTemp           As ADODB.Recordset
        Dim LngItemID        As Long
        Dim LngUnitID        As Long
        Dim ColorID          As Integer
        Dim sizeid           As Integer
        Dim ClassId          As Integer
        Dim ParrtNoCode      As String
        Dim ItemDetailedCode As String
        Dim StrSerial        As String
        Dim rsDummy          As New ADODB.Recordset
        Dim mQty             As Double
        CodeFromBarCode = TxtItemCodeB1
        Dim QtyFromFrmSa5 As Double

        If UCase(frmname) = "FRMSALEBILL5" And val(TxtQuantity) <> 0 Then
            QtyFromFrmSa5 = IIf(val(TxtQuantity) <> 0, val(TxtQuantity), 1)
        Else
            QtyFromBarCode = 1
        End If

        Dim mIsBarCodeUnit        As Boolean
        Dim DateFromBarCode       As String
        Dim SupplierFromBarCode   As String
        Dim CODEQty               As String
        Dim ExpiryDateFromBarCode As Variant
        Dim mFromNo               As Double
     
        '#1234101182456
        CodeFromBarCode = TxtItemCodeB1
    
        If Len(TxtItemCodeB1) >= 18 Then
            'TxtItemCodeB1 = mId(TxtItemCodeB1, 2, Len(TxtItemCodeB1) - 1)
            CodeFromBarCode = mId(TxtItemCodeB1, 1, 4)
    
            '
            DateFromBarCode = mId(TxtItemCodeB1, 5, 6)
            ExpiryDateFromBarCode = mId(DateFromBarCode, 1, 2) & "/" & mId(DateFromBarCode, 3, 2) & "/20" & mId(DateFromBarCode, 5, 2)
            SupplierFromBarCode = mId(TxtItemCodeB1, 11, 3)
            CODEQty = mId(TxtItemCodeB1, 14, 5)
    
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                TxtQuantity.Text = CODEQty
            End If

            TxtItemCodeB1 = CodeFromBarCode
      
        End If

        If SystemOptions.IsBarCodeByUnit Then
            StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitID DefaultUnit"
            StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
            StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
            StrSQL = StrSQL & " WHERE     (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
            StrSQL = StrSQL & "   dbo.TblItemsUnits.barCodeNo2 ='" & TxtItemCodeB1 & "'   )"
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp.EOF Then
        
                StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
                StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
                StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
                StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1) AND (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
                StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & TxtItemCodeB1 & "'   ) and (  TblItems.IsArchive =0) "
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            Else
                mIsBarCodeUnit = True
            End If

        Else
            StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
            StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
            StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
            StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
            StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & TxtItemCodeB1 & "'   )"
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        End If
    
        If RsTemp.EOF Then
            GetQtyCodeFormBarCode

            If SystemOptions.IsBarCodeByUnit Then
                StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitID DefaultUnit"
                StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
                StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
                StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & CodeFromBarCode & "'  OR"
                StrSQL = StrSQL & "   dbo.TblItemsUnits.barCodeNo2 ='" & CodeFromBarCode & "'  Or dbo.TblItems.barCodeNO ='" & val(CodeFromBarCode) & "' )"
                RsTemp.Close
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.EOF Then
                    StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
                    StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
                    StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
                    StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & CodeFromBarCode & "'  OR"
                    StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & CodeFromBarCode & "'  Or dbo.TblItems.barCodeNO ='" & val(CodeFromBarCode) & "' )"
                    RsTemp.Close
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                Else
                    mIsBarCodeUnit = True
                End If
            End If
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

                If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = GridTransType.internalorder Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Then
             
                    TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
             
                    If SystemOptions.workWithCustomerContract = True And Me.GridTrans = InvoiceTransaction Then
                        TxtPrice.Text = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)

                        If TxtPrice.Text = 0 Then
                            TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
                        End If
                    End If

                    StrSQL = " Select HaveWeight from TblItems"
                    StrSQL = " SELECT tu.UnitName,tu.UnitNamee, ISNULL(HaveWeight, 0) AS IsWeight FROM TblUnites AS tu Where UnitId = " & LngUnitID
                    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Dim mQtyString As String

                    If Not rsDummy.EOF Then
                        If (Trim(rsDummy!UnitName & "") = "Õ»…" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PEICE" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PP" Or rsDummy("IsWeight").value = 0) And mIsBarCodeUnit And Trim(rsDummy!UnitName & "") <> "ﬂ— Ê‰ ﬂ»Ì—" Then
                    
                            mQtyString = left(CStr(QtyFromBarCode), 2)
                            mQty = val(mQtyString)
                        ElseIf mIsBarCodeUnit And (rsDummy("IsWeight").value = 1 Or Trim(rsDummy!UnitName & "") <> "Õ»…" Or UCase(Trim(rsDummy!UnitNamee & "")) <> "PEICE" Or Trim(rsDummy!UnitName & "") = "ﬂ— Ê‰ ﬂ»Ì—") Then
                            mQty = QtyFromBarCode / 1000
                        Else
                            mQtyString = CODEQty
                            mQty = val(mQtyString)
                        End If
                    End If
             
                ElseIf Me.GridTrans = Destruction Or Me.GridTrans = InventoryOut Or Me.GridTrans = NewGard Or m_GridTrans = MoveItems Or m_GridTrans = StockSettlement Then

                    If Not TxtSerial Is Nothing Then
                        StrSerial = Trim$(TxtSerial.Text)
                    End If
 
                    '        unitid = val(.Cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                    TxtPrice.Text = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), LngUnitID, val(Me.StoreName.BoundText))
                 
                End If
            
                If SystemOptions.IsDogeMode = True Then
                    TxtQuantity.Text = ""
                Else
                    TxtQuantity.Text = IIf(mQty = 0, 1, TxtQuantity.Text)
                End If

                If UCase(frmname) = "FRMSALEBILL5" And val(QtyFromFrmSa5) <> 0 Then
                    TxtQuantity.Text = val(QtyFromFrmSa5)
                Else
                
                End If
             
                TxtQuantity.Tag = "F"
                TxtQuantity.Enabled = True
                '    Me.Grid.Rows = Me.Grid.Rows + 1
        
                With Grid

                    If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                        .rows = .rows + 1
                    End If
         
                    If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Or Me.GridTrans = GridTransType.internalorder Or Me.GridTrans = MoveItems Or Me.GridTrans = ProductionOrder Or Me.GridTrans = PurchaseTransaction Then
                        If val(TxtQuantity.Text) = 0 Then
        
                            If SystemOptions.IsDogeMode = True Then
                                TxtQuantity.Text = ""
                            Else
                                TxtQuantity.Text = IIf(Trim(QtyFromBarCode) = "", 1, QtyFromBarCode)
                            End If
                        End If

                        AddItemInGrid .rows - 1, val(DCboItemName.BoundText), val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode, , , , , CStr(ExpiryDateFromBarCode)
                        TxtQuantity.Text = ""

                    End If

                    FrmCustomerDisplay.LblInformation.Caption = DCboItemName.Text & vbNewLine & "  Price   " & val(Me.TxtPrice.Text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
                    FrmCustomerDisplay.lblitemid = LngItemID
                End With
           
                TxtQuantity.Text = 0
                DCboItemCode.BoundText = 0
                Me.DCboItemName.Text = ""
 
                'DCboItemCode_KeyDown vbKeyReturn, 0
                Me.TxtItemCodeB1.Text = ""
                'Unload FrmItemSearch2
                Me.TxtItemCodeB1.SetFocus
         
            Else
        
                Exit Sub
            End If

        Else
            Dim Msg

            If SystemOptions.UserInterface = ArabicInterface Then
                'MsgBox "·«ÌÊÃœ ’‰› »Â–« «·ﬂÊœ", vbCritical
                Msg = "·«ÌÊÃœ ’‰› »Â–« «·ﬂÊœ"
            Else
                'MsgBox "There is no item in this code", vbCritical
                Msg = "There is no item in this code"
 
            End If

            If Me.GridTrans = InvoiceTransaction And Not (txtItemCodeSearch2 Is Nothing) Then
                MsgBox Msg, vbCritical
            End If

            'messageResult = MsgBoxPause(, msg, App.title, vbExclamation, MsssageSeconde)
                
            TxtQuantity.Text = ""
            TxtItemCodeB1.Text = ""
 
            Me.TxtItemCodeB1.SetFocus
        End If

    End If

End Sub
 


 
 

Public Sub TxtItemCodeB1_KeyDown1(KeyCode As Integer, Shift As Integer)
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
Dim mQtyBarFromSales5 As Double
'WaelNew
If UCase(frmname) = "FRMSALEBILL5" And val(TxtQuantity) <> 0 Then
    QtyFromBarCode = val(TxtQuantity)
    mQtyBarFromSales5 = val(TxtQuantity)
Else
    QtyFromBarCode = 1
End If
   
            
            
Dim mIsBarCodeUnit As Boolean
If SystemOptions.IsBarCodeByUnit Then
         StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitID DefaultUnit"
    StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
    StrSQL = StrSQL & " WHERE     (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
    StrSQL = StrSQL & "   dbo.TblItemsUnits.barCodeNo2 ='" & TxtItemCodeB1 & "'   )"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If RsTemp.EOF Then
        
        StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
        StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
        StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1) AND (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
        StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & TxtItemCodeB1 & "'   ) and (  TblItems.IsArchive =0) "
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Else
        mIsBarCodeUnit = True
    End If
Else
    StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
    StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
    StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & TxtItemCodeB1 & "'  OR"
    StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & TxtItemCodeB1 & "'   )"
     Set RsTemp = New ADODB.Recordset
     RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
End If
    
If RsTemp.EOF Then
    GetQtyCodeFormBarCode
    If SystemOptions.IsBarCodeByUnit Then
             StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.UnitID DefaultUnit"
        StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
        StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & CodeFromBarCode & "'  OR"
        StrSQL = StrSQL & "   dbo.TblItemsUnits.barCodeNo2 ='" & CodeFromBarCode & "'  Or dbo.TblItems.barCodeNO ='" & val(CodeFromBarCode) & "' )"
        RsTemp.Close
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp.EOF Then
            StrSQL = "SELECT     dbo.TblItems.Fullcode, dbo.TblItems.barCodeNO, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit"
            StrSQL = StrSQL & "  FROM         dbo.TblItems INNER JOIN"
            StrSQL = StrSQL & "   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
            StrSQL = StrSQL & " WHERE     (dbo.TblItemsUnits.DefaultUnit = 1)  and (  TblItems.IsArchive =0)  AND (dbo.TblItems.Fullcode ='" & CodeFromBarCode & "'  OR"
            StrSQL = StrSQL & "   dbo.TblItems.barCodeNO ='" & CodeFromBarCode & "'  Or dbo.TblItems.barCodeNO ='" & val(CodeFromBarCode) & "' )"
            RsTemp.Close
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Else
            mIsBarCodeUnit = True
        End If
    End If
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
         If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ShowPrice Or Me.GridTrans = Template Or Me.GridTrans = ReturnSalling Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = GridTransType.internalorder Then
             TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))
             
             StrSQL = " SELECT tu.UnitName,tu.UnitNamee FROM TblUnites AS tu Where UnitId = " & LngUnitID
             rsDummy.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Not rsDummy.EOF Then
                If Trim(rsDummy!UnitName & "") = "Õ»…" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PEICE" Or UCase(Trim(rsDummy!UnitNamee & "")) = "PP" Or mIsBarCodeUnit Then
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

             
                         
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                TxtQuantity.Text = mQty
            End If

            TxtQuantity.Tag = "F"
            TxtQuantity.Enabled = True
        '    Me.Grid.Rows = Me.Grid.Rows + 1
        
    With Grid

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If
         
        If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = ProductionOrder Then
            If UCase(frmname) = "FRMSALEBILL5" And val(mQtyBarFromSales5) <> 0 Then
                mQty = val(mQtyBarFromSales5)

            End If
             
                         
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                TxtQuantity.Text = mQty
            End If

             
                   AddItemInGrid .rows - 1, val(DCboItemName.BoundText), val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode
        End If
FrmCustomerDisplay.LblInformation.Caption = DCboItemName.Text & vbNewLine & "  Price   " & val(Me.TxtPrice.Text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
   FrmCustomerDisplay.lblitemid = LngItemID
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
 DCboItemCode.SetFocus
 
    
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
             TxtPrice.Text = GetItemPrice(LngItemID, 1, LngUnitID, , CheckCustomerSaleType(val(Customer.BoundText)))

  
            If SystemOptions.IsDogeMode = True Then
                TxtQuantity.Text = ""
            Else
                TxtQuantity.Text = 1
            End If
            TxtQuantity.Tag = "F"
            TxtQuantity.Enabled = True
        '    Me.Grid.Rows = Me.Grid.Rows + 1
        
    With Grid

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If
                   AddItemInGrid .rows - 1, val(DCboItemName.BoundText), val(TxtQuantity.Text), val(Me.TxtPrice.Text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode
FrmCustomerDisplay.LblInformation.Caption = DCboItemName.Text & vbNewLine & "  Price   " & val(Me.TxtPrice.Text) & vbNewLine & " —Ì‹«· ”⁄ÊœÌ SAR "
   FrmCustomerDisplay.lblitemid = LngItemID
     End With
           
           
 
         'DCboItemCode_KeyDown vbKeyReturn, 0
          Me.TxtItemCodeB.Text = ""
         ' Unload FrmItemSearch2
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
            If Me.GrdTBar.Buttons.count = 0 Then
            
            
                SetToolBar
            End If
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
    If m_GridTrans2 <> 0 Then
        frmname = Me.Grid.Parent.Name & m_GridTrans2
    Else
        frmname = Me.Grid.Parent.Name
    End If
    StrSetting = GetSetting(SystemOptions.SysRegsAppPath, "Interface Setting" & "\" & user_id, frmname, "")

If frmname = "frmsalebill5" Then
 
 StrSetting = " Ser-0;Code-1;Name-2;PrintName-3;Count-4;AddSerials-5;Price-6;DeliveredQty2-7;FoxyNo-8;ItemCase-9;HaveSerial-10;Serial-11;guaranteeTime-12;DiscountType-13;DiscountVal-14;Valu-15;order_no-16;OrderArrivalDate-17;ItemType-18;ItemCostPrice-19;GranteeType-20;GranteeStartDate-21;GranteeEndDate-22;PofTransID-23;RegularMaintenancedates-24;RegularMaintenanceIDS-25;ItemProfit-26;LSPH-27;RSPH-28;LCYL-29;RCYL-30;LAXIS-31;RAXIS-32;LIPD-33;RIPD-34;LADD-35;RADD-36;LSH-37;RSH-38;LPRISM-39;RPRISM-40;ItemsDetailsNewidea-41;"
StrSetting = StrSetting & "  discountvalue-42;UnitID-43;ItemSize-44;ColorID-45;UnitPrice-46;ClassId-47;StoreIDLoc-48;StoreID2-49;CorrectionID-50;ProductionDate-51;"
StrSetting = StrSetting & " GranteeData-52;GoldData-53;ExpiryDate-54;LotNO-55;SupplierID-56;catalog-57;countris-58;ScurrencyID-59;ParrtNoCode-60;"
StrSetting = StrSetting & " ItemDetailedCode-61;Remarks-62;QtyFaqtors-63;CostPriceTk-64;Height-65;OUTR-66;ForUnit-67;MethodCalc-68;lowering-69;increase-70;"
StrSetting = StrSetting & " INR-71;length-72;ItemQty-73;Width-74;Area-75;NoCount-76;LastPurchasePrice-77;ItemBalance-78;ItemSalesYear-79;New_Rolling-80;ReminRequ-81;"
StrSetting = StrSetting & " IsExpirDate-82;L-83;W-84;H1-85;H2-86;Vatyo-87;Vat-88;TypeVAT-89;MixNo-90;MaxQty-91;MaxUnitID-92;CusID2-93;NoHours-94;PriceNoHours-95;TotalPriceNoHours-96;"
StrSetting = StrSetting & "  TotalInvoiceQty-97;ISSUEDQTY-98;ManufactCompany-99;CountryOrg-100;EmpID4-101;SalesValue-102;TotalPrice-103"


'StrSetting = StrSetting & "  LSPH-27;RSPH-28;LCYL-29;RCYL-30;LAXIS-31;RAXIS-32;LIPD-33;RIPD-34;LADD-35;RADD-36;LSH-37;RSH-38;LPRISM-39;RPRISM-40;ItemsDetailsNewidea-41;discountvalue-42;UnitID-43;ItemSize-44;ColorID-45;UnitPrice-46;ClassId-47;StoreIDLoc-48;StoreID2-49;CorrectionID-50;ProductionDate-51;GranteeData-52;GoldData-53;ExpiryDate-54;LotNO-55;SupplierID-56;catalog-57;countris-58;ScurrencyID-59"
'StrSetting = StrSetting & "   ;ParrtNoCode-60;ItemDetailedCode-61;Remarks-62;QtyFaqtors-63;CostPriceTk-64;Height-65;OUTR-66;ForUnit-67;MethodCalc-68;lowering-69;increase-70;INR-71;length-72;ItemQty-73;Width-74;Area-75;NoCount-76;"
'StrSetting = StrSetting & "  LastPurchasePrice-77;ItemBalance-78;ItemSalesYear-79;New_Rolling-80;ReminRequ-8"
'StrSetting = StrSetting & "  1;IsExpirDate-82;L-83;W-84;H1-85;H2-86;Vatyo-87;Vat-88;TypeVAT-89;MixNo-90;MaxQty-91;MaxUnitID-92;CusID2-93;NoHours-94;PriceNoHours-95;TotalPriceNoHours-96;TotalInvoiceQty-97;ISSUEDQTY-98;ManufactCompany-99;CountryOrg-100;EmpID4-101;SalesValue-102;TotalPrice-103"



End If
    If StrSetting = "" Then
        Exit Sub
    End If

    VarCols = Split(StrSetting, ";", , vbTextCompare)

    If UBound(VarCols) > 0 Then

        For i = 0 To UBound(VarCols)
            VarColSet = Empty
            VarColSet = Split(CStr(VarCols(i)), "-", , vbTextCompare)

            With Grid
                On Error Resume Next
            '    .ColPosition(.ColIndex(CStr(VarColSet(0)))) = CLng(VarColSet(1))
                
            End With

        Next i

    End If

    StrShowSet = GetSetting(SystemOptions.SysRegsAppPath, "Cols Setting" & "\" & user_id, frmname, "")

If frmname = "frmsalebill5" Then
  StrShowSet = "Ser-True;Code-False;Name-False;PrintName-True;Count-False;AddSerials-True;Price-False;DeliveredQty2-True;FoxyNo-True;ItemCase-True;HaveSerial-True;Serial-True;guaranteeTime-True;DiscountType-False;DiscountVal-False;Valu-False;order_no-True;OrderArrivalDate-True;ItemType-True;ItemCostPrice-True;GranteeType-True;GranteeStartDate-True;GranteeEndDate-True;PofTransID-True;RegularMaintenancedates-True;RegularMaintenanceIDS-True;ItemProfit-True;LSPH-True;RSPH-True;"
StrShowSet = StrShowSet & "LCYL-True;RCYL-True;LAXIS-True;RAXIS-True;LIPD-True;RIPD-True;LADD-True;RADD-True;LSH-True;RSH-True;LPRISM-True;RPRISM-True;ItemsDetailsNewidea-True;discountvalue-True;UnitID-True;ItemSize-True;ColorID-True;UnitPrice-True;ClassId-True;StoreIDLoc-True;StoreID2-True;CorrectionID-True;ProductionDate-True;GranteeData-True;GoldData-True;ExpiryDate-True;LotNO-True;SupplierID-True;catalog-True;countris-True;ScurrencyID-True;ParrtNoCode-True;"
StrShowSet = StrShowSet & "ItemDetailedCode-True;Remarks-True;QtyFaqtors-True;CostPriceTk-True;Height-True;OUTR-True;ForUnit-True;MethodCalc-True;lowering-True;increase-True;"
StrShowSet = StrShowSet & "INR-True;length-True;ItemQty-True;Width-True;Area-True;NoCount-True;LastPurchasePrice-True;ItemBalance-True;ItemSalesYear-True;New_Rolling-True;ReminRequ-True;"
StrShowSet = StrShowSet & "IsExpirDate-True;L-True;W-True;H1-True;H2-True;Vatyo-True;Vat-True;TypeVAT-True;MixNo-True;MaxQty-True;MaxUnitID-True;CusID2-True;NoHours-True;"
StrShowSet = StrShowSet & "PriceNoHours-True;TotalPriceNoHours-True;TotalInvoiceQty-True;ISSUEDQTY-True;ManufactCompany-True;CountryOrg-True;EmpID4-True;"
StrShowSet = StrShowSet & "SalesValue-True;TotalPrice-True"


 
End If


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
        .rows = .FixedRows

        '-----------------
                If Me.Grid.ColIndex("NewItem") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "’‰› ÃœÌœ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "New Item"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "NewItem"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NewItem"))
        End If
        
        
        If Me.Grid.ColIndex("Ser") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„”·”· «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Ser"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Ser"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Ser"))
        End If


'        If Me.Grid.ColIndex("project") <> -1 Then
'            .rows = .rows + 1
'            i = i + 1
'            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
'
'                            If SystemOptions.UserInterface = ArabicInterface Then
'                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„‘—Ê⁄"
'                            Else
'                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "project"
'                            End If
'            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "project"
'            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("project"))
'          End If
          
        '-----------
       If m_GridTrans2 = 2 Then
           If Me.Grid.ColIndex("PumpName") <> -1 Then
                .rows = .rows + 1
                i = i + 1
                .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
    
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„÷Œ…"
                Else
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Pump Name"
                End If
                .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "PumpName"
                .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PumpName"))
            End If
        End If
        
        
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·’‰›"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Code"
        End If
        
        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Code"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Code"))
        '------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«”„ «·’‰›"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Name  "
        End If
          
        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Name"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Name"))


        '------------
        '«·ÊÕœ…
        If Me.Grid.ColIndex("UnitID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ÊÕœ… «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Unit"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "UnitID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("UnitID"))
        End If
        
        '«··Ê‰
        If Me.Grid.ColIndex("ColorID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "·Ê‰ «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Item Color  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ColorID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ColorID"))
        End If

        '---------------------
        '«·„ﬁ«”
        If Me.Grid.ColIndex("ItemSize") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„ﬁ«” «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Size"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemSize"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemSize"))
        End If

        '-------------------
        '«·›∆…
        If Me.Grid.ColIndex("ClassId") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "›—“ «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Item Category  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ClassId"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ClassId"))
        End If

        '„Ê«ﬁ⁄ «·Õ›Ÿ
        If Me.Grid.ColIndex("StoreIDLoc") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„Êﬁ⁄ «· Œ“Ì‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Store Location"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "StoreIDLoc"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StoreIDLoc"))
        End If

        '-------------------
                    If Me.Grid.ColIndex("StoreID2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„Œ“‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Store"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "StoreID2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StoreID2"))
        End If


    '********************************
        '«·«Ã—«¡«  «· ’ÕÌÕÌ…
        If Me.Grid.ColIndex("CorrectionID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·«Ã—«¡«  «· ’ÕÌÕÌ…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Corrective action"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CorrectionID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CorrectionID"))
        End If

        '-------------------
    
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì…"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Quantity"
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Count"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Count"))
        '-------------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·”⁄—"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Price"
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Price"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Price"))
        '-------------------
        
        '-------------------
        If Me.Grid.ColIndex("DiscountType") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·Œ’„"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Discount Type"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "DiscountType"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountType"))
        End If

        '-------------------
        If Me.Grid.ColIndex("DiscountVal") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·Œ’„"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Discount Value"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "DiscountVal"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DiscountVal"))
        End If

        '-------------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬁÌ„…"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Value"
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Valu"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Valu"))
        '-------------------

        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Item Status  "
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemCase"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCase"))
        '-------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "·Â ”Ì—Ì«·"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Have Serial"
        End If
        
        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "HaveSerial"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("HaveSerial"))
        '---------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "≈÷«›… «·”Ì—Ì«·"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Add Serials"
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "AddSerials"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AddSerials"))
        '-------------------
        .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·”Ì—Ì«·"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  Serial NO"
        End If

        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Serial"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Serial"))
        '-------------------


 
        
        
        If Me.Grid.ColIndex("guaranteeTime") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„œ… «·÷„«‰ »«·‘Â—"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "guarantee Time"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "guaranteeTime"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("guaranteeTime"))
        End If
        
        If Me.Grid.ColIndex("ItemType") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·’‰› "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ItemType "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemType"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemType"))
        End If
        
        If Me.Grid.ColIndex("order_no") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·ÿ·»ÌÂ"
            Else
            
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "order no"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "order_no"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("order_no"))
        End If
        
       
        If Me.Grid.ColIndex("OrderArrivalDate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "   «—ÌŒ Ê’Ê· «·‘Õ‰…   "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Order Arrival Date  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "OrderArrivalDate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OrderArrivalDate"))
        End If
        
        If Me.Grid.ColIndex("LineShahn") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„’«—Ì› «·‘Õ‰"
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LineShahn"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LineShahn"))
        End If
        
        If Me.Grid.ColIndex("ItemCostPrice") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— «· ﬂ·›…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ItemCostPrice"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemCostPrice"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCostPrice"))
        End If
        
        
        If Me.Grid.ColIndex("CostPriceTk") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— «· ﬂ·›… ··«” ·«„"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "CostPriceTk"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CostPriceTk"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CostPriceTk"))
        End If
        
'CostPriceTk

        If Me.Grid.ColIndex("PofTransID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ „Œ’’"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "PofTransID"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "PofTransID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PofTransID"))
        End If
        
        If Me.Grid.ColIndex("ItemProfit") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—»Õ «·’ ›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ItemProfit"
            
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemProfit"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemProfit"))
        End If
      
        If Me.Grid.ColIndex("FoxyNo") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·»—‰«„Ã"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Prog No"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "FoxyNo"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("FoxyNo"))
        End If
                                        
        If Me.Grid.ColIndex("Account_Serial") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
            
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂÊœ  «·Õ”«»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Account Serial"

            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Account_Serial"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Account_Serial"))
        End If
        
        
        
         
         If Me.Grid.ColIndex("Account_Serial2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
            
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·Õ”«»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Account Serial"

            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Account_Serial2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Account_Serial2"))
        End If
        
        If Me.Grid.ColIndex("Account_Code") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·Õ”«»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂÊœ   2"

            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Account_Code"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = False ' Not Me.Grid.ColHidden(Me.Grid.ColIndex("Account_Code"))
        End If
        
        If Me.Grid.ColIndex("Account_Name") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «”„ «·Õ”«»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Account_Name"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Account_Name"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Account_Name"))
        End If
        
        If Me.Grid.ColIndex("AccountCode") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "code 2"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂÊœ   2"

            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "AccountCode"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AccountCode"))
        End If
        
        If Me.Grid.ColIndex("AccountName") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «”„ «·Õ”«»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " AccountName"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "AccountName"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AccountName"))
        End If
        
        If Me.Grid.ColIndex("GardQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «·ﬂ„ÌÂ «·œ› —ÌÂ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Ac QTy  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GardQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GardQty"))
        End If
        
        If Me.Grid.ColIndex("Gardresult") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·›—ﬁ  "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Different  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Gardresult"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult"))
        End If
        
        If Me.Grid.ColIndex("Gardresult1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then

                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄Ã“  "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Minus  "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Gardresult1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult1"))
        End If
        
        If Me.Grid.ColIndex("Gardresult2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·“Ì«œ…  "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Increase"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Gardresult2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Gardresult2"))
        End If
        
        If Me.Grid.ColIndex("ProductionDate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «—ÌŒ   «·«‰ «Ã"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Production Date"
            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ProductionDate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ProductionDate"))
        End If

        If Me.Grid.ColIndex("GranteeData") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "»Ì«‰«  «·÷„«‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Gurantee Date"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GranteeData"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GranteeData"))
        End If
        
                If Me.Grid.ColIndex("GoldData") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ›«’Ì· «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Details"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GoldData"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GoldData"))
        End If
        
        
        
        
                  If Me.Grid.ColIndex("ParrtNoCode") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬂÊœ «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " part Code"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ParrtNoCode"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ParrtNoCode"))
        End If
        
        
                          If Me.Grid.ColIndex("ItemDetailedCode") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬂÊœ   Õ·Ì·Ì «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Detailed part Code"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemDetailedCode"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemDetailedCode"))
        End If
        
   
   
           
                          If Me.Grid.ColIndex("PrintName") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«”„ «·’‰› «·„»«⁄"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Print Name"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "PrintName"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PrintName"))
        End If
        
     
        'ItemCostPrice
        
        
        If Me.Grid.ColIndex("ExpiryDate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «—ÌŒ   «·«‰ Â«¡"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Expire Date"
            End If
           
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ExpiryDate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExpiryDate"))
        End If
        
        
                       If Me.Grid.ColIndex("IsExpirDate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «—ÌŒ ÌœÊÌ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Date Manual"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "IsExpirDate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("IsExpirDate"))
        End If




        'LotNO
        If Me.Grid.ColIndex("LotNO") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ «··Êÿ"
 
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LOT NO"

            End If
            
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LotNO"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LotNO"))
        End If
        
        If Me.Grid.ColIndex("DamageshowQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·Â«·ﬂ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Damage"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "DamageshowQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DamageshowQty"))
        End If
          
        If Me.Grid.ColIndex("GranteeData") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "»Ì«‰«  «·÷„«‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Gurantee Date"
            End If
           
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GranteeData"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GranteeData"))
        End If

       If Me.Grid.ColIndex("GoldData") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ›«’Ì· «·ﬁÿ⁄Â"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Details"
            End If
           
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GoldData"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GoldData"))
        End If

       If Me.Grid.ColIndex("ShowQty2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Ê“‰ «·–Â»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Gold Weight"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ShowQty2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowQty2"))
        End If
      
      
             If Me.Grid.ColIndex("showPrice2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— «·Ã—«„"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Gram Value"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "showPrice2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("showPrice2"))
        End If
        
        
                     If Me.Grid.ColIndex("Value2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·–Â»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Gold Value"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Value2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Value2"))
        End If
        
        
        
      

              If Me.Grid.ColIndex("ShowQty1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Ê“‰ «·«·„«”"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Diamond Weight"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ShowQty1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowQty1"))
        End If
        
        
                If Me.Grid.ColIndex("showPrice1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— «·ﬁÌ—«ÿ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Carat Price"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "showPrice1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("showPrice1"))
        End If
        
        
              
                If Me.Grid.ColIndex("Value1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·√·„«”"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Diamond  Total"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Value1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Value1"))
        End If
        
        '*********************************
                 '⁄„·… «·«·„«”
                    If Me.Grid.ColIndex("ScurrencyID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄„·…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Currency"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ScurrencyID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ScurrencyID"))
        End If
        '*********************************
        
              '*********************************
                 '„⁄œ· ⁄„·… «·«·„«”
                    If Me.Grid.ColIndex("Scurrenyrate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„⁄œ·"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Rate"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Scurrenyrate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Scurrenyrate"))
        End If
        '*********************************
        
                      '*********************************
                 '  —ﬁ„ «·«„— ⁄·Ï „”‰ÊÏ «·”ÿ— ›Ì ”‰œ «” ·«„ «‰ «Ã  «„
                    If Me.Grid.ColIndex("NProductionOrderNO") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ «„— «·«‰ «Ã"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Production Order  NO"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "NProductionOrderNO"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NProductionOrderNO"))
        End If
        '*********************************
        
        
        
     '//////////////////////////////////////////////////////////////////////////////////////////
             '*********************************
        
                      '*********************************
                 '  ÿ·» «·‘—«¡ «·œ«Œ·Ì
                    If Me.Grid.ColIndex("LastPurchaseDate") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «—ÌŒ «Œ— ‘—«¡"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Last Purchase Date"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LastPurchaseDate"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchaseDate"))
        End If
        
       
                    If Me.Grid.ColIndex("LastPurchasePrice") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«Œ—  ”⁄—   ‘—«¡"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Last Purchase Price"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LastPurchasePrice"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchasePrice"))
        End If
               
               
               
               
                    If Me.Grid.ColIndex("LastPurchaseqty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂ„Ì… «Œ— ‘—«¡"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Last Purchase Qty"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LastPurchaseqty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LastPurchaseqty"))
        End If
               
               
                                 If Me.Grid.ColIndex("ItemBalance") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—’Ìœ «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Balance "
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemBalance"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemBalance"))
        End If
        
        
           If Me.Grid.ColIndex("DeliveredQty2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·„”·„…  "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Delivered Qty"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "DeliveredQty2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DeliveredQty2"))
        End If
        
        
        
                   If Me.Grid.ColIndex("RecivePeriod") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„œ… «· Ê—Ìœ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " Recive Period"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RecivePeriod"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RecivePeriod"))
        End If
        
        
               
                    If Me.Grid.ColIndex("AverageIssue") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„ Ê”ÿ «·«”‰Â·«ﬂ «·‘Â—Ì"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "íMonthly Average"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "AverageIssue"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AverageIssue"))
        End If
        
                            If Me.Grid.ColIndex("AverageIssueyraly") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·«”‰Â·«ﬂ «·”‰ÊÌ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Yearly Average"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "AverageIssueyraly"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("AverageIssueyraly"))
        End If
        
        
        
                          If Me.Grid.ColIndex("RequestLimit") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Õœ «⁄«œ… «·ÿ·»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "íRequest Limit"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RequestLimit"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RequestLimit"))
        End If
        
       '//////////////////////////////////////////////////////////////////////////////////////////
       
     '*********************
     
        
                          If Me.Grid.ColIndex("Shipping") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·‘Õ‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Shipping"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Shipping"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Shipping"))
        End If
     
                          If Me.Grid.ColIndex("countris") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«· ’‰Ì⁄"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Made In"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "countris"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("countris"))
        End If
        
     
     '***********************
        
             
        '***************************
       If Me.Grid.ColIndex("catalog") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "   «·ﬂ «·ÊÃ "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " catalog"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "catalog"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("catalog"))
        End If
        '**************************
        
        
        
        
        
        
        
             If Me.Grid.ColIndex("Salaries1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«ÃÊ— ’Ì«€Â"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Formulation fees"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Salaries1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Salaries1"))
        End If
        
        
                
             If Me.Grid.ColIndex("Salaries2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«ÃÊ—  —ﬂÌ»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Installation fees"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Salaries2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Salaries2"))
        End If
        
        
          If Me.Grid.ColIndex("TotalCost") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·«Ã„«·Ì «·⁄«„"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Total Cost"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalCost"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalCost"))
        End If
        
        
        '---------------------
        
        
        '  ««·„Ê—œÌ‰
        If Me.Grid.ColIndex("SupplierID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„Ê—œÌ‰"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Supplier Name"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "SupplierID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SupplierID"))
        End If

        '-------------------
        
       '-------------------
        If Me.Grid.ColIndex("ExtraType") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·⁄„Ê·…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "commision Type"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ExtraType"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExtraType"))
        End If

        '-------------------
        If Me.Grid.ColIndex("ExtraVal") <> -1 Then '
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·⁄„Ê·…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "commision Value"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ExtraVal"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ExtraVal"))
        End If

 


        '-------------------
                If Me.Grid.ColIndex("SBillNO") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "—ﬁ„ «·›« Ê—…"
            Else
            
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Bill no"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "SBillNO"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SBillNO"))
        End If
        
        
                         If Me.Grid.ColIndex("ItemCostPrice") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«· ﬂ·›…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Cost Price"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemCostPrice"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemCostPrice"))
        End If
            
         

                     
        
         '??????
   
             ' Exit Sub
          
        
        
  '*****************
              If Me.Grid.ColIndex("Height") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·«— ›«⁄"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Height"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Height"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Height"))
        End If
        ''''«·⁄—÷
             If Me.Grid.ColIndex("Width") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄—÷"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Width"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Width"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Width"))
        End If
           
       '
         
               ''''«·„”«Õ…
             If Me.Grid.ColIndex("Area") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„”«Õ…"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Area"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Area"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Area"))
          End If
          
                        '«·”„ﬂ
             If Me.Grid.ColIndex("thickness") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·”„ﬂ"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "thickness"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "thickness"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("thickness"))
          End If
 
          
                     ''''«·⁄œœ
             If Me.Grid.ColIndex("NoCount") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄œœ"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "No"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "NoCount"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NoCount"))
        End If

'*****************
        ''//////////  «Ê  ÕÃ„ „»Ì⁄«  «·⁄«„
             If Me.Grid.ColIndex("ItemSalesYear") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Status"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ItemSalesYear"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ItemSalesYear"))
        End If




             If Me.Grid.ColIndex("Remarks") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„·«ÕŸ« "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Remarks"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Remarks"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Remarks"))
        End If





   '*****************
        ''//////////ÃœÌœ «Ê „ œ«Ê·
             If Me.Grid.ColIndex("New_Rolling") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Õ«·… «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Item Status"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "New_Rolling"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("New_Rolling"))
        End If


        
                        '„ »ﬁÌ Õœ «·ÿ·»
        If Me.Grid.ColIndex("ReminRequ") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "„ »ﬁÌ Õœ «·ÿ·»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Remaining Request"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ReminRequ"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ReminRequ"))
        End If


                   If Me.Grid.ColIndex("Select") <> -1 Then
             .rows = .rows + 1
        i = i + 1
        .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ÕœÌœ"
        Else
            .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Select"
        End If
        
        .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Select"
        .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Select"))
        
        End If

'--------------------------------------------------------------------

                  If Me.Grid.ColIndex("L") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "L"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("L"))
        End If
                '''/////////////
          If Me.Grid.ColIndex("TotalPrice") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·”⁄— «·‘«„·"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Total"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalPrice"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalPrice"))
        End If
        
        '''/////////////
          If Me.Grid.ColIndex("W") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "W"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "W"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "W"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("W"))
        End If
         If Me.Grid.ColIndex("H1") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "H1"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "H1"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "H1"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("H1"))
        End If
      ''///
            If Me.Grid.ColIndex("H2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "H2"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "H2"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "H2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("H2"))
        End If
        '--------------------------------------------------------------------
                            If Me.Grid.ColIndex("Vatyo") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰”»… «· VAT % "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "VAT %"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Vatyo"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Vatyo"))
        End If
        
        
        
                    If Me.Grid.ColIndex("Vat") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «· VAT"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "VAT  "
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Vat"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Vat"))
        End If
        
        
                
        If Me.Grid.ColIndex("TypeVAT") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «· VAT "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "VAT Type"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TypeVAT"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TypeVAT"))
        End If
        
        
             If Me.Grid.ColIndex("MintName") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ‰Ê⁄ «·’Ì«‰…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Maintenance"
            End If
           
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "MintName"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MintName"))
          End If
                If Me.Grid.ColIndex("GroupMint") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "›∆… «·’Ì«‰…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Group"
            End If
           
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "GroupMint"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("GroupMint"))
        End If


         If Me.Grid.ColIndex("QtyFaqtors") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄·«ﬁ… „⁄ «·ÊÕœ…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Qty Factor"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "QtyFaqtors"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("QtyFaqtors"))
        End If


'*********************************************
If Me.Grid.ColIndex("MixNo") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂÊœ «·„ﬂ”"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Mix. No"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "MixNo"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MixNo"))
        End If

'*********************************************

'*******************

          If Me.Grid.ColIndex("MaxUnitID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·ÊÕœ… «·ﬂ»—Ï"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Max Unit"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "MaxUnitID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MaxUnitID"))
        End If
       ''//////////
          If Me.Grid.ColIndex("MaxQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·ﬂ„Ì… «·ﬂ»—Ï"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Max Qty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "MaxQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MaxQty"))
        End If

'****************
   If Me.Grid.ColIndex("CusID2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «”„ «·⁄„Ì·"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Customer Name"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CusID2"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CusID2"))
        End If
'***************************************************************
       If Me.Grid.ColIndex("NoHours") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "⁄œœ «·”«⁄« "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "No. Hours"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "NoHours"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NoHours"))
        End If
        
         If Me.Grid.ColIndex("PriceNoHours") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ﬂ·›… «·”«⁄… "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Cost Hour"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "PriceNoHours"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PriceNoHours"))
        End If
        
          If Me.Grid.ColIndex("TotalPriceNoHours") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " «Ã„«·Ì  ﬂ·›… «·”«⁄«  "
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Total Cost Hours"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalPriceNoHours"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalPriceNoHours"))
        End If
 
          If Me.Grid.ColIndex("TotalInvoiceQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·«Ã„«·Ì…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Total Û QTY"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalInvoiceQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalInvoiceQty"))
        End If



 




 ShowGridColFrm1 (i)
        
    End With


End Sub
Sub ShowGridColFrm1(i As Integer)
  With FrmGridColsShow.FG
  
'**************************
          If Me.Grid.ColIndex("ISSUEDQTY") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·„‰’—› ”«»ﬁ«"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ISSUED Û QTY"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ISSUEDQTY"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ISSUEDQTY"))
        End If



          If Me.Grid.ColIndex("TotalQtyIn") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalQtyIn"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalQtyIn"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalQtyIn"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalQtyIn"))
        End If



          If Me.Grid.ColIndex("TotalBalance") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalBalance"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalBalance"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalBalance"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalBalance"))
        End If

        If Me.Grid.ColIndex("Select") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Select"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Select"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Select"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Select"))
          End If
          

        If Me.Grid.ColIndex("project") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·„‘—Ê⁄"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "project"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "project"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("project"))
          End If
          
          
            If Me.Grid.ColIndex("pand") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·»‰œ"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "pand"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "pand"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("pand"))
          End If
          
            If Me.Grid.ColIndex("opera") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·⁄„·Ì…"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "opera"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "opera"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("opera"))
          End If
          


            If Me.Grid.ColIndex("ProfitType") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·—»ÕÌ…"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·—»ÕÌ…"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ProfitType"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ProfitType"))
          End If
   
   
   
       If Me.Grid.ColIndex("TotalBalance") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalBalance"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalBalance"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalBalance"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalBalance"))
        End If
        
        
         If Me.Grid.ColIndex("ProfitValue") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·—»ÕÌ…"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÌ„… «·—»ÕÌ…"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ProfitValue"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ProfitValue"))
          End If
          
          
         If Me.Grid.ColIndex("NetProfit") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "  «·«Ã„«·Ì „⁄ «·—»Õ"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·«Ã„«·Ì „⁄ «·—»Õ  "
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "NetProfit"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("NetProfit"))
          End If
         '////////////
                 '
           If Me.Grid.ColIndex("OldID") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "OldID"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "OldID"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "OldID"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OldID"))
          End If
          
         
           If Me.Grid.ColIndex("OrderLineId") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "OrderLineId"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "OrderLineId"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "OrderLineId"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OrderLineId"))
          End If
           
          
          
         '///////////
        '«·ÿÊ·
           If Me.Grid.ColIndex("length") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ÿÊ·"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "length"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "length"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("length"))
          End If
       
       
               ''''ﬁÿ— Œ«—ÃÌ
             If Me.Grid.ColIndex("OUTR") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÿ— Œ«—ÃÌ"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "OUTR"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "OUTR"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("OUTR"))
          End If
       
       
       
   


       '***************************
        
     '**************************
       
               ''''ﬁÿ— œ«Œ·Ì
             If Me.Grid.ColIndex("INR") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "ﬁÿ— œ«Œ·Ì"
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "INR"
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "INR"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("INR"))
          End If
       
       '***************************
      
       '**************************
       
               ''''  ⁄—÷ «·„—›ﬁ« 
             If Me.Grid.ColIndex("ShowAttatch2") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
                
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " ⁄—÷ «·„—›ﬁ«    "
                            Else
                                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Show Attatch "
                            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "INR"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ShowAttatch2"))
          End If
       
       '***************************
           
           
  
  
            If Me.Grid.ColIndex("ManufactCompany") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·‘—ﬂ… «·’«‰⁄…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Manufacturing Company"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ManufactCompany"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ManufactCompany"))
        End If
        
           If Me.Grid.ColIndex("CountryOrg") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = " »·œ «·„‰‘√"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Country of Origin"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CountryOrg"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CountryOrg"))
        End If
  
  
          If Me.Grid.ColIndex("EmpID4") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "   «”„ «·„ÊŸ›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Employee Name"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "EmpID4"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("EmpID4"))
        End If
        
  
  
           If Me.Grid.ColIndex("SalesValue") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— »Ì⁄ «·’‰›"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Sales Value"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "SalesValue"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("SalesValue"))
        End If



  
           If Me.Grid.ColIndex("UnitPrice") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "”⁄— «·ÊÕœ…  " ' «·«”„ ⁄—»Ì
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Unit price" ' «·«”„ «‰Ã·Ì“Ì
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "UnitPrice"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("UnitPrice"))
        End If

'            If .ColIndex("ProfitType") <> -1 Then
'                .TextMatrix(0, .ColIndex("ProfitType")) = "ProfitType"
'
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    .ColComboList(.ColIndex("ProfitType")) = "#1;·«ÌÊÃœ |#2;—»ÕÌ… »ﬁÌ„…|#3;—»ÕÌ… »‰”»…|#"
'                ElseIf SystemOptions.UserInterface = EnglishInterface Then
'                    .ColComboList(.ColIndex("ProfitType")) = "#1;NO |#2;Value|#3;Percent profit|#"
'                End If
'
'            End If


'***************************************************GLASSES'***************************************************
            If .ColIndex("ItemBalance2") <> -1 Then
                .TextMatrix(0, .ColIndex("ItemBalance2")) = "«·—’Ìœ ›Ï «·„Œ“‰ «·„ÿ·Ê» «· ÕÊÌ· „‰Â"
                
            Else
        
            End If

            If .ColIndex("RequestTypeNo") <> -1 Then
                .TextMatrix(0, .ColIndex("RequestTypeNo")) = "RequestTypeNo"
            Else
        
            End If

            If .ColIndex("RequestTypeName") <> -1 Then
                .TextMatrix(0, .ColIndex("RequestTypeName")) = "‰Ê⁄ «·ÿ·»"
                .ColComboList(.ColIndex("RequestTypeName")) = "#1;ÿ·» ‘—«¡|#2;ÿ·»  ÕÊÌ· „Œ“‰Ì"
            Else
        
            End If



            If .ColIndex("StoreIDAvi") <> -1 Then
                .TextMatrix(0, .ColIndex("StoreIDAvi")) = "StoreIDAvi"
            Else
        
            End If

            If .ColIndex("StoreIDAviName") <> -1 Then
                .TextMatrix(0, .ColIndex("StoreIDAviName")) = "«·„Œ«“‰ «·„ «Õ ›ÌÂ« "
            Else
        
            End If

        If Me.Grid.ColIndex("RequestTypeName") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "‰Ê⁄ «·ÿ·»"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Request TypeName"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RequestTypeName"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RequestTypeName"))
        End If
        
                    
                
        If Me.Grid.ColIndex("RCYL") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RCYL"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RCYL"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RCYL"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RCYL"))
        End If
        
        
        
        
                     

        
                
                         If Me.Grid.ColIndex("RSPH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RSPH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RSPH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RSPH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RSPH"))
        End If
   

        If Me.Grid.ColIndex("LIPD") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LIPD"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LIPD"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LIPD"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LIPD"))
        End If
        
            If Me.Grid.ColIndex("RIPD") <> -1 Then
                .rows = .rows + 1
                i = i + 1
                .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RIPD"
                Else
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RIPD"
                End If

                .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RIPD"
                .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RIPD"))
            End If
        
        
            If Me.Grid.ColIndex("LCL") <> -1 Then
                .rows = .rows + 1
                i = i + 1
                .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LCL"
                Else
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LCL"
                End If

                .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LCL"
                .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LCL"))
            End If
        
       
            If Me.Grid.ColIndex("RCL") <> -1 Then
                .rows = .rows + 1
                i = i + 1
                .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RCL"
                Else
                    .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RCL"
                End If

                .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RCL"
                .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RCL"))
            End If
        
         
        
                  If Me.Grid.ColIndex("LADD") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LADD"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LADD"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LADD"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LADD"))
        End If
        
                  If Me.Grid.ColIndex("RADD") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RADD"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RADD"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RADD"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RADD"))
        End If
        
                  If Me.Grid.ColIndex("LSH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LSH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LSH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LSH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LSH"))
        End If
        
         
                           If Me.Grid.ColIndex("RSH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RSH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RSH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RSH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RSH"))
        End If
        
        
                                If Me.Grid.ColIndex("LPRISM") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LPRISM"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LPRISM"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LPRISM"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LPRISM"))
        End If
        
        
                                If Me.Grid.ColIndex("RPRISM") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RPRISM"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RPRISM"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RPRISM"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RPRISM"))
        End If
        
                             If Me.Grid.ColIndex("LAXIS") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LAXIS"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LAXIS"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LAXIS"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LAXIS"))
        End If
        
        
        
        
                          If Me.Grid.ColIndex("LCYL") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LCYL"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LCYL"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LCYL"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LCYL"))
        End If
         
         

        
                          If Me.Grid.ColIndex("LSPH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LSPH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "LSPH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LSPH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LSPH"))
        End If
        
 
 
        
        If Me.Grid.ColIndex("ReadSPH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "R Read SPH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "R Read SPH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ReadSPH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ReadSPH"))
        End If
        
   
        If Me.Grid.ColIndex("ReadCYL") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "R ReadCYL"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "R ReadCYL"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "ReadCYL"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ReadCYL"))
        End If
        
        
'        If Me.Grid.ColIndex("ReadSPH") <> -1 Then
'            .Rows = .Rows + 1
'            i = i + 1
'            .TextMatrix(.Rows - 1, .ColIndex("ColIndex")) = i
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ReadSPH"
'            Else
'                .TextMatrix(.Rows - 1, .ColIndex("ColName")) = "ReadSPH"
'            End If
'
'            .TextMatrix(.Rows - 1, .ColIndex("ColKey")) = "ReadSPH"
'            .TextMatrix(.Rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("ReadSPH"))
'        End If
'
 
 
    
        If Me.Grid.ColIndex("LReadCYL") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L ReadCYL"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L ReadCYL"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LReadCYL"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LReadCYL"))
        End If
        
        
        If Me.Grid.ColIndex("LReadSPH") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L ReadSPH"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "L ReadSPH"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "LReadSPH"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("LReadSPH"))
        End If
        
        
                          If Me.Grid.ColIndex("RAXIS") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RAXIS"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "RAXIS"
            End If

            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "RAXIS"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("RAXIS"))
        End If
                
              
              If Me.Grid.ColIndex("TotalQtyIn") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalQtyIn"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "TotalQtyIn"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "TotalQtyIn"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("TotalQtyIn"))
        End If



              
        

        If Me.Grid.ColIndex("PrevQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬁ—«¡… «·”«»ﬁ…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Prev"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "PrevQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("PrevQty"))
        End If


        If Me.Grid.ColIndex("CurrentQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬁ—«¡… «·Õ«·Ì…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Prev"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CurrentQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CurrentQty"))
        End If

  

        If Me.Grid.ColIndex("UsedQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·„” Œœ„…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Prev"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "UsedQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("UsedQty"))
        End If

  
        If Me.Grid.ColIndex("StillQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "«·ﬂ„Ì… «·„ »ﬁÌ…"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Prev"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "StillQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StillQty"))
        End If

    If Me.Grid.ColIndex("StillPumbQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "StillPumbQty"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "StillPumbQty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "StillPumbQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("StillPumbQty"))
        End If

    If Me.Grid.ColIndex("CashQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "CashQty"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "CashQty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "CashQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("CashQty"))
        End If
        
        If Me.Grid.ColIndex("Cash") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Cash"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Cash"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Cash"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Cash"))
        End If
        
        
               
        If Me.Grid.ColIndex("MadaQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "MadaQty"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "MadaQty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "MadaQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("MadaQty"))
        End If
         
        
        If Me.Grid.ColIndex("Mada") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Mada"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Mada"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Mada"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Mada"))
        End If
        
                   
        If Me.Grid.ColIndex("VisaQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "VisaQty"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "VisaQty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "VisaQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("VisaQty"))
        End If
        
        
                
        If Me.Grid.ColIndex("Visa") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Visa"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Visa"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Visa"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Visa"))
        End If
        
                            
        If Me.Grid.ColIndex("DeferredQty") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "DeferredQty"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "DeferredQty"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "DeferredQty"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("DeferredQty"))
        End If
                        
        If Me.Grid.ColIndex("Deferred") <> -1 Then
            .rows = .rows + 1
            i = i + 1
            .TextMatrix(.rows - 1, .ColIndex("ColIndex")) = i

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Deferred"
            Else
                .TextMatrix(.rows - 1, .ColIndex("ColName")) = "Deferred"
            End If
            .TextMatrix(.rows - 1, .ColIndex("ColKey")) = "Deferred"
            .TextMatrix(.rows - 1, .ColIndex("ColShow")) = Not Me.Grid.ColHidden(Me.Grid.ColIndex("Deferred"))
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
    If m_GridTrans2 <> 0 Then
        frmname = Me.Grid.Parent.Name & m_GridTrans2
    Else
        frmname = Me.Grid.Parent.Name
    End If
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


Private Sub txtPrice_Change()
'If SystemOptions.CanChangePriceUpOnly Then
'                 Dim rsDummyPrice As ADODB.Recordset
'                 Dim StrSQL  As String
'                 Set rsDummyPrice = New ADODB.Recordset
'
'                    StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID,TBLiTEMSUnits.UnitSalesPrice, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
'
'                    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
'                    StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId "
'                    'and DefaultUnit = 1"
'                    StrSQL = StrSQL + " Where TblItems.ItemID=" & val(DCboItemName.BoundText)
'                    StrSQL = StrSQL + "  and TBLiTEMSUnits.DefaultUnit=1"
'                    rsDummyPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
'                    If Not rsDummyPrice.EOF Then
'                        If val(TxtPrice) < val(rsDummyPrice!UnitSalesPrice & "") Then
'                             TxtPrice = val(rsDummyPrice!UnitSalesPrice & "")
'                        End If
'                    End If
'
'                End If

End Sub

Private Sub TxtPrice_GotFocus()


    TxtPrice.backcolor = vbYellow
    SelectText TxtPrice

    If Me.DCboItemName.BoundText <> "" Then
        If m_ShowQtyToolTip = True Then
         
           If m_GridTrans2 <> InvoiceTransactionCompose Then
                ShowItemTip ItemPriceToolTip, val(Me.DCboItemName.BoundText), 0, 0, , False, Me.TxtPrice.hWnd
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
       If m_GridTrans2 = InvoiceTransaction Then
                Dim rsDummy As New ADODB.Recordset
                Dim StrSQL  As String
                StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                
                StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId and DefaultUnit = 1"
                StrSQL = StrSQL + " Where TblItems.ItemID=" & val(Me.DCboItemName.BoundText)
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    If val(rsDummy!MinSelingPrice & "") > 0 And val(TxtPrice) < val(rsDummy!MinSelingPrice & "") Then
                        MsgBox "·« Ì„ﬂ‰ «·»Ì⁄ »«ﬁ· „‰ «·Õœ «·«œ‰Ï ··”⁄—"
                        TxtPrice.SetFocus
                        Exit Sub
                    ElseIf val(rsDummy!MaxSelingPrice & "") > 0 And val(TxtPrice) > val(rsDummy!MaxSelingPrice & "") Then
                        MsgBox "·« Ì„ﬂ‰ «·»Ì⁄ »«⁄·Ï  „‰ «·Õœ «·«⁄·Ï ··”⁄—"
                        TxtPrice.SetFocus
                        Exit Sub
                    End If
                    
                End If
            End If
            
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

Private Sub txtPrice_Validate(Cancel As Boolean)
If Me.GridTrans = purchaseorder Or Me.GridTrans = purchaseorderrequest Or Me.GridTrans = PurchaseTransaction Then
    Exit Sub
End If
If SystemOptions.CanChangePriceUpOnly Then
                 Dim rsDummyPrice As ADODB.Recordset
                 Dim StrSQL  As String
                 Set rsDummyPrice = New ADODB.Recordset

                    StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID,TBLiTEMSUnits.UnitSalesPrice, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "

                    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                    StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId "
                    'and DefaultUnit = 1"
                    StrSQL = StrSQL + " Where TblItems.ItemID=" & val(DCboItemName.BoundText)
                    StrSQL = StrSQL + "  and TBLiTEMSUnits.DefaultUnit=1"
                    rsDummyPrice.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                    If Not rsDummyPrice.EOF Then
                        If val(TxtPrice) < val(rsDummyPrice!UnitSalesPrice & "") Then
                             TxtPrice = val(rsDummyPrice!UnitSalesPrice & "")
                        End If
                    End If

                End If
End Sub

Private Sub TxtQuantity_GotFocus()
    TxtQuantity.backcolor = vbYellow
    TxtQuantity.SetFocus
                TxtQuantity.SelStart = 0
            TxtQuantity.SelLength = Len(TxtQuantity.Text)
            
End Sub

Private Sub TxtQuantity_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
   
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtQuantity.Text, 0)
End Sub

 
 Private Sub TxtQuantity_LostFocus()
  
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
             MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             

             
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
            'check customer price val(Customer.BoundText)
            
                .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID, Purchase, CheckCustomerSaleType(val(Customer.BoundText)))
            

                If SystemOptions.workWithCustomerContract = True And (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = 3) Then
        
                    .TextMatrix(Row, .ColIndex("Price")) = getpricebycustomrContract(val(Customer.BoundText), LngUnitID, LngItemID, , val(DcboEmp.BoundText), TxtCashCustomerName, Me.DtpBillDate.value)
                    If .TextMatrix(Row, .ColIndex("Price")) = 0 Then
                        .TextMatrix(Row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(Row, .ColIndex("Count"))), LngUnitID, Purchase)
                    End If
                 
                 
                End If
                
       
                
                If Me.GridTrans = Returntransaction Then
                
             
                    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
                
                    FromDate = FirstPeriodDateInthisYear
                    GetlastPurchasedata 20, (LngItemID), FromDate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
                
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

    If Me.Grid.rows = Me.Grid.FixedRows Then
        Exit Function
    End If

    With Me.Grid
        If Me.Grid.rows = LngRow Then
            Exit Function
        End If
        LngItemID = val(Me.Grid.TextMatrix(LngRow, Grid.ColIndex("Code")))
        LngUnitID = val(.cell(flexcpData, LngRow, .ColIndex("UnitID")))

        If LngItemID <> 0 Then

            Select Case Me.Grid.cell(flexcpChecked, LngRow, Me.Grid.ColIndex("HaveSerial"))

                Case flexChecked
                    IntHaveSerial = 1

                Case flexUnchecked
                    IntHaveSerial = 2
            End Select

            StrItemSerial = Trim(.TextMatrix(LngRow, .ColIndex("Serial")))

            If Not (IntHaveSerial = 1 And StrItemSerial = "") Then
           If Me.GridTrans <> InvoiceTransaction And Me.GridTrans <> Returntransaction Then
            Exit Function
           End If
                If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
                    'DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID, ModernWeightAverage, , , , , LngUnitID)
                    
                    DblItemCost = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrItemSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(Me.TxtNots.Text), LngUnitID, val(Me.StoreName.BoundText))
                
                Else
                    
                    DblItemCost = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrItemSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(Me.TxtNots.Text), LngUnitID, val(Me.StoreName.BoundText))
                    '  DblItemCost = GetCostItemPrice(LngItemID, IntHaveSerial, StrItemSerial, StrTransID)
                End If
          
            
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemCostPrice")) = DblItemCost
            
                If Me.Grid.ColIndex("ItemProfit") = -1 Then
                    Exit Function
                End If
                DblTotalItemCost = (DblItemCost * val(Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("Count"))))
                
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("ItemProfit")) = Round(val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) - DblTotalItemCost, SystemOptions.SysDefCurrencyForamt)
            
                
                '------------------------------------------------------------------
           If SystemOptions.showcostColorininvoice = True Then
                If val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) = DblTotalItemCost Then
                    Me.Grid.cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbYellow
                ElseIf val(.TextMatrix(LngRow, .ColIndex("Valu"))) * val(txt_Currency_rate.Text) < DblTotalItemCost Then
                    Me.Grid.cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbRed
                Else
                    Me.Grid.cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = 0
                End If
            Else
                     Me.Grid.cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = 0
            End If
            
                '------------------------------------------------------------------
            
                Me.Grid.TextMatrix(LngRow, Me.Grid.ColIndex("PofTransID")) = StrTransID

                If Not LblInvProfit Is Nothing Then
                    DblInvProfit = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemProfit"), .rows - 1, .ColIndex("ItemProfit"))
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
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer, Optional StoreID As Integer, Optional lotNo As String, Optional ExpiryDate As String, Optional ProductionDate As String)

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
    ElseIf LngRow > Me.Grid.rows - 1 Then
        Me.Grid.rows = LngRow + 1
    End If

    If Not (RsItems.BOF Or RsItems.EOF) Then
        If RsItems("ItemID").value <> LngItemID Then
            RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
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
                        RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
                    End If
                End If

                If RsItems("RelatedItem").value = True Then
                    Set Frm = New FrmAddItemAttachs
                    Frm.LoadItemData LngItemID, DblQty
                    Frm.show vbModal

                    If Frm.UserCancelled = True Then
                        Unload Frm
                    Else

                        For j = Frm.FG.FixedRows To Frm.FG.rows - 1

                            If Frm.FG.cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
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
             
         .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
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
 
                    UnitID = val(.cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                    .TextMatrix(LngRow, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), UnitID, val(Me.StoreName.BoundText))
                
                
                ElseIf Me.GridTrans = ProductionOrder Then
                    .TextMatrix(LngRow, .ColIndex("Price")) = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
          
                
            ElseIf Me.GridTrans = internalorder Then
                
                
               
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    FromDate = FirstPeriodDateInthisYear
 
    
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Or Grid.ColIndex("LastPurchaseqty") <> -1 Or Grid.ColIndex("LastPurchasePrice") <> -1 Or Grid.ColIndex("RequestLimit") <> -1 Then
                           
                           GetlastPurchasedata 20, (LngItemID), FromDate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
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
                            GetMonthlyAverage 19, (LngItemID), FromDate, Date, AverageMonthly
                               .TextMatrix(LngRow, .ColIndex("AverageIssue")) = AverageMonthly
                           End If
                           
                           
                                   If Grid.ColIndex("AverageIssueyraly") <> -1 Then
                            GetYearlyAverage 19, (LngItemID), FromDate, Date, AverageIssueyraly
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
             
         .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
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
                                            .cell(flexcpPicture, LngRow, .ColIndex("Ser")) = _
                                                 mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
                    Else

                        If .ColIndex("Ser") <> -1 Then
                            .cell(flexcpPicture, LngRow, .ColIndex("Ser")) = Nothing
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
                                                                 .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
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
                                         .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = discId
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
                .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignRightCenter
            Else
                .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignLeftCenter
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

                    For j = Frm.FG.FixedRows To Frm.FG.rows - 1

                        If Frm.FG.cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
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
                    .cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").value
                    longunitid = RsUnitData("UnitID").value
                    
           If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").value
          Else
                   .TextMatrix(LngRow, .ColIndex("UnitID")) = IIf(IsNull(RsUnitData("UnitNamee").value), RsUnitData("UnitName").value, RsUnitData("UnitNamee").value)
          End If
                Else
                    .cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
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
        If subcall = False Then
            Exit Function
        End If
        
        
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
For i = 1 To .rows - 1
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
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer, Optional StoreID As Integer, Optional Qty2 As Double = 0)
    
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
        My_SQL = My_SQL & " and (dbo.Transactions.StoreID = " & StoreID & ")"
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
If val(.TextMatrix(LngRow, .ColIndex("IsExpirDate"))) = 1 Then
    Exit Sub
End If
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
            .rows = .rows + 1
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
                    .rows = .rows + 1
                     AddItemInGrid2 .rows - 1, LngItemID, TempQty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, StoreID, lotNo, ExpiryDate, ProductionDate
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
                           .rows = .rows + 1
                           
        'End If
               
                     AddItemInGrid2 .rows - 1, LngItemID, Qty, SngItemPrice, IntItemCase, StrItemSerial, FromHeader, DefalutUnitID, DefalutUnitName, freeitemsqty, subcall, ColorID, itemsize, ClassId, ParrtNoCode, ItemDetailedCode, TransactionDate, discountype, StoreID
                  ReCalVAT LngRow
         End If
End With
                    rs.Close
   End If
   
     ColorExperDate
     SentTypeVAT
End Sub
 
   

'
  Public Function AddItemInGrid(LngRow As Long, _
                              LngItemID As Long, _
                              Optional DblQty As Double = 1, _
                              Optional SngItemPrice As Single = 0, _
                              Optional IntItemCase As Integer = 0, _
                              Optional StrItemSerial As String = "", _
                              Optional FromHeader As Boolean = False, _
                              Optional DefalutUnitID As Long = 0, _
                              Optional DefalutUnitName As String, Optional freeitemsqty As Integer = 0, _
                              Optional subcall As Boolean = True, Optional ColorID As Integer = 1, Optional itemsize As Integer = 1, Optional ClassId As Integer = 1, Optional ParrtNoCode As String, Optional ItemDetailedCode As String, Optional TransactionDate As Variant, Optional discountype As Integer, Optional ByVal mEmpId As Long, Optional ByVal IsEyeData As Integer = 0, Optional ExpiryDate As String, Optional haveoffer As Boolean)

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
           
                 ''''''
        If val(TxtQuantity) = 0 Then TxtQuantity = 1
    If Me.GridTrans = NewGard And SystemOptions.itemsWorkWithDates = False Then
        Dim LngFindRow As Long
        LngFindRow = Grid.FindRow(LngItemID, 0, Grid.ColIndex("Code"), False, True)
        If LngFindRow <> -1 Then
            LngRow = LngFindRow
            DblQty = DblQty + val(Grid.TextMatrix(LngRow, Grid.ColIndex("Count")))
        End If
    
    End If
    
'
'    If Not RSPH Is Nothing And (DCboItemCode.Text = "" And DCboItemName.Text = "") Then
'            If FramePay.Visible = False Then
'                Dim s As String
'                Dim rsDummy As New ADODB.Recordset
'                s = " SELECT ti.ItemCode,ti.ItemID,ti.ItemName,ti.SphereID,ti.CylinderID,g.GroupName,tu.UnitName,SPHTable.SPHT,CLYTable.CLYT"
'                s = s & "      FROM TblItems AS ti INNER JOIN Groups AS g ON g.GroupID = ti.GroupID"
'                s = s & " INNER JOIN TblItemsUnits ON TblItemsUnits.ItemID = ti.ItemID"
'                s = s & " INNER JOIN TblUnites AS tu ON tu.UnitID = TblItemsUnits.UnitID"
'                s = s & " INNER JOIN SPHTable ON ti.SphereID = SPHTable.ID"
'                s = s & " INNER JOIN CLYTable ON ti.CylinderID = CLYTable.ID"
'                s = s & " Where  (SPHTable.ID = " & val(GetSphCylID(0, RSPH) & " and CLYTable.ID= " & val(GetSphCylID(1, RCLY))) & ")"
'                s = s & " Or (SPHTable.ID = " & val(GetSphCylID(0, LSPH)) & " and CLYTable.ID= " & val(GetSphCylID(1, LCLY)) & ")"
'
'                loadgrid s, GrdItems, True, False
'                If GrdItems.Rows > 1 Then
'                    FramePay.Visible = True
'                    'Exit Sub
'
'                End If
'
'        End If
'    End If
            
    '------------------  Â‰« «·«÷«›…
    GetDefaultItemUnit LngItemID, LngUnitID
    If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ProductionOrder) And SystemOptions.AllowItemByRow = True Or ((Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN) And SystemOptions.AllowItemByRowOut = True) Or (Me.GridTrans = MoveItems And SystemOptions.AllowItemByRowMove = True) Then
                               
                               
    If CheckMaultyItem(LngRow, LngItemID, DefalutUnitID, DblQty) = False Then
    If LngRow = 0 Then
        LngRow = Me.NewEmptyRow
    ElseIf LngRow > Me.Grid.rows - 1 Then
        Me.Grid.rows = LngRow + 1
    End If
    Else
    Me.Calculate LngRow, , , True
End If
Else
   If LngRow = 0 Then
        LngRow = Me.NewEmptyRow
    ElseIf LngRow > Me.Grid.rows - 1 Then
        Me.Grid.rows = LngRow + 1
    End If
End If
    If Not (RsItems.BOF Or RsItems.EOF) Then
        If RsItems("ItemID").value <> LngItemID Then
            RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
        End If
    End If

    
  '      GetDefaultItemUnit LngItemID, LngUnitID
    
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
                        RsItems.Find "ItemID=" & LngItemID & "", , adSearchForward, 1
                    End If
                End If

                If RsItems("RelatedItem").value = True Then
                    Set Frm = New FrmAddItemAttachs
                    Frm.LoadItemData LngItemID, DblQty
                    Frm.show vbModal

                    If Frm.UserCancelled = True Then
                        Unload Frm
                    Else

                        For j = Frm.FG.FixedRows To Frm.FG.rows - 1

                            If Frm.FG.cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
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
            If mEmpId <> 0 Then
                .TextMatrix(LngRow, .ColIndex("EmpID4")) = mEmpId
            End If
            .TextMatrix(LngRow, .ColIndex("Code")) = LngItemID
            .TextMatrix(LngRow, .ColIndex("Name")) = LngItemID
            If mPumpId <> 0 Then
               ' .TextMatrix(LngRow, .ColIndex("PumpName")) = mPumpId
            End If
            If UCase(frmname) = "FRMSALEBILL6" And m_GridTrans2 = 2 Then
                .TextMatrix(LngRow, .ColIndex("PrevQty")) = txtPumbDet1
                .TextMatrix(LngRow, .ColIndex("CurrentQty")) = txtPumbDet2
                .TextMatrix(LngRow, .ColIndex("PumpName")) = cmbPumbType.BoundText
                .TextMatrix(LngRow, .ColIndex("StillPumbQty")) = txtPumbDet1
                s = "Select * from tblPumpType where id = " & val(cmbPumbType.BoundText)
                
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                        .TextMatrix(LngRow, .ColIndex("AmountH")) = IIf(IsNull(rsDummy("AmountH")), "", Trim(rsDummy("AmountH").value))
                        '.TextMatrix(LngRow, .ColIndex("AmountHComm")) = IIf(IsNull(RsDetails("AmountHComm")), "", Trim(RsDetails("AmountHComm").value))
                        '.TextMatrix(LngRow, .ColIndex("DetailsPump")) = IIf(IsNull(RsDetails("DetailsPump")), "", Trim(RsDetails("DetailsPump").value))
                        .TextMatrix(LngRow, .ColIndex("Account_CodeComm")) = IIf(IsNull(rsDummy("Account_CodeComm")), "", Trim(rsDummy("Account_CodeComm").value))
                        .TextMatrix(LngRow, .ColIndex("Account_Code")) = IIf(IsNull(rsDummy("Account_Code")), "", Trim(rsDummy("Account_Code").value))
                        .TextMatrix(LngRow, .ColIndex("IsOther")) = IIf(IsNull(rsDummy("IsOther")), 0, Trim(rsDummy("IsOther").value))
                        .TextMatrix(LngRow, .ColIndex("StoreID2")) = IIf(IsNull(rsDummy("StoreID")), 0, Trim(rsDummy("StoreID").value))
           

                End If
            End If
            
            .cell(flexcpData, LngRow, .ColIndex("UnitID")) = LngUnitID
            If .ColIndex("chkTaxExempt") <> -1 Then
            If val(RsItems!ItemWithOutVAT & "") = 0 Then
                .TextMatrix(LngRow, .ColIndex("chkTaxExempt")) = False
            Else
            
                .TextMatrix(LngRow, .ColIndex("chkTaxExempt")) = IIf((RsItems!ItemWithOutVAT & ""), True, False)
            End If
            End If
            
            .TextMatrix(LngRow, .ColIndex("ExpiryDate")) = ExpiryDate
            
            
            If .ColIndex("Account_Name") <> -1 Then
                
                  .TextMatrix(LngRow, .ColIndex("Account_Code")) = cmbAccounts.BoundText
                    '.TextMatrix(row, .ColIndex("Account_Serial2")) = IIf(IsNull(rsDummy("Account_Serial").Value), "", rsDummy("Account_Serial").Value)

                    
                .TextMatrix(LngRow, .ColIndex("Account_Name")) = cmbAccounts.Text
                   
            End If
            If IsEyeData = 1 Then
            
                If Not RSPH Is Nothing Then
    
                .TextMatrix(LngRow, .ColIndex("RSPH")) = RSPH.Text
                .TextMatrix(LngRow, .ColIndex("RCYl")) = RCLY.Text
                .TextMatrix(LngRow, .ColIndex("RAxis")) = RAxis.Text
                
                .TextMatrix(LngRow, .ColIndex("LSPH")) = LSPH.Text
                .TextMatrix(LngRow, .ColIndex("LCYL")) = LCLY.Text
                .TextMatrix(LngRow, .ColIndex("LAxis")) = LAxis.Text
                

                
              '  val (RsDetails("RSPH") & "") + 0.5 * val(RsDetails("RCYL") & "")
                .TextMatrix(LngRow, .ColIndex("RCL")) = val(RSPH) + 0.5 * val(RCLY)  'txtRCL.Text
                .TextMatrix(LngRow, .ColIndex("LCL")) = val(LSPH) + 0.5 * val(LCLY)
                
                .TextMatrix(LngRow, .ColIndex("RIPD")) = RIPD.Text
                .TextMatrix(LngRow, .ColIndex("LIPD")) = LIPD.Text
                .TextMatrix(LngRow, .ColIndex("RADD")) = RADD.Text
                .TextMatrix(LngRow, .ColIndex("LADD")) = LADD.Text
                .TextMatrix(LngRow, .ColIndex("RSH")) = RSH.Text
                .TextMatrix(LngRow, .ColIndex("LSH")) = LSH.Text
                .TextMatrix(LngRow, .ColIndex("RPRISM")) = RPRISM.Text
                .TextMatrix(LngRow, .ColIndex("LPRISM")) = LPRISM.Text
                .TextMatrix(LngRow, .ColIndex("ItemID2")) = mMainRow
                
'                .TextMatrix(LngRow, .ColIndex("LCYL")) = LCLY.Text
                
                End If
            End If
            .TextMatrix(LngRow, .ColIndex("StoreIDLoc")) = GetLocationID(val(.TextMatrix(LngRow, .ColIndex("Code"))))
            .TextMatrix(LngRow, .ColIndex("HaveSerial")) = RsItems("HaveSerial").value
            If Me.GridTrans = PurchaseTransaction Or Me.GridTrans = Returntransaction Then
               If SystemOptions.LinkSupplerWithItem = True Then
                If GetSuppler(LngItemID) <> 0 Then
                    Customer.BoundText = GetSuppler(LngItemID)
                End If
               
               End If
            
            End If
            
                              
            If Me.GridTrans = InvoiceTransaction Or Me.GridTrans = PurchaseTransaction Or Me.GridTrans = ReturnSalling Or Me.GridTrans = Returntransaction Or Me.GridTrans = InventoryOut Or Me.GridTrans = INVENTORYIN Then
                If SystemOptions.IsGeometricProportions Then
                    LngUnitID = val(.cell(flexcpData, LngRow, .ColIndex("UnitID")))
                   
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
                'SalimSalim SalimSalim
'            DblQty = IIf(SystemOptions.DefaultQtyTrans = 0, "", SystemOptions.DefaultQtyTrans)

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
             
         .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
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
 
                    UnitID = val(.cell(flexcpData, LngRow, .ColIndex("UnitID")))
                
                    .TextMatrix(LngRow, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, StrSerial, , SystemOptions.SysMainStockCostMethod, , , DtpBillDate, val(TxtInvID.Text), UnitID, val(Me.StoreName.BoundText))
                
                
                ElseIf Me.GridTrans = ProductionOrder Then
                    .TextMatrix(LngRow, .ColIndex("Price")) = GET_COST_PRICE_FOR_PRODUCT_ITEM(LngItemID)
          
                
            ElseIf Me.GridTrans = internalorder Or Me.GridTrans = purchaserequest Then
                
                
               
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

    FromDate = FirstPeriodDateInthisYear
 
    
                           If Grid.ColIndex("LastPurchaseDate") <> -1 Or Grid.ColIndex("LastPurchaseqty") <> -1 Or Grid.ColIndex("LastPurchasePrice") <> -1 Or Grid.ColIndex("RequestLimit") <> -1 Then
                           
                           GetlastPurchasedata 20, (LngItemID), FromDate, Date, LastPurchaseDate, LastPurchasePrice, LastPurchaseqty
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
                            GetMonthlyAverage 19, (LngItemID), FromDate, Date, AverageMonthly
                               .TextMatrix(LngRow, .ColIndex("AverageIssue")) = AverageMonthly
                           End If
                           
                           
                                   If Grid.ColIndex("AverageIssueyraly") <> -1 Then
                            GetYearlyAverage 19, (LngItemID), FromDate, Date, AverageIssueyraly
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
             
         .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
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
                                            .cell(flexcpPicture, LngRow, .ColIndex("Ser")) = _
                                                 mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
                    Else

                        If .ColIndex("Ser") <> -1 Then
                            .cell(flexcpPicture, LngRow, .ColIndex("Ser")) = Nothing
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
 If ExpiryDate = "" Then
             .TextMatrix(LngRow, .ColIndex("ExpiryDate")) = ""
 End If
             .TextMatrix(LngRow, .ColIndex("ProductionDate")) = ""
            .TextMatrix(LngRow, .ColIndex("IsExpirDate")) = ""

            
                          '«·’‰› «·„·Õﬁ €«·»« „Ã«‰Ì
                        If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnSalling) And discountype = 4 Then
                                If .ColIndex("DiscountType") <> -1 Then
                                        If SystemOptions.attacheditemsisfree = True Or haveoffer = True Then '22
                                           .TextMatrix(LngRow, .ColIndex("Price")) = 0
                                                                 .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 4
                                                                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 4
                                                   
                                        End If
                                End If
                                
                        End If
                        
'⁄—Ê÷ «·«’‰«›
''//////ShowPrice
                        
ShowPricess LngRow, haveoffer, discountype, LngItemID

         If Grid.ColIndex("ParrtNoCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ParrtNoCode")) = ParrtNoCode
                           End If
                      
                      
                      If Grid.ColIndex("ItemDetailedCode") <> -1 Then
                        
                               .TextMatrix(LngRow, .ColIndex("ItemDetailedCode")) = ItemDetailedCode
                           End If
                           
                           
            '-----------------------------------
            If SystemOptions.UserInterface = ArabicInterface Then
                .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignRightCenter
            Else
                .cell(flexcpPictureAlignment, 0, 0, .rows - 1, .Cols - 1) = flexPicAlignLeftCenter
            End If

            '----------------------------------
            '--------≈÷«›… «·√’‰«› «·„·Õﬁ… »«·’‰›
            If RsItems("RelatedItem").value = True Then
                Set Frm = New FrmAddItemAttachs
                Frm.Txtqty = DblQty
                Frm.LoadItemData LngItemID, DblQty
                Frm.show vbModal

                If Frm.UserCancelled = True Then
                    Unload Frm
                Else

                    For j = Frm.FG.FixedRows To Frm.FG.rows - 1

                        If Frm.FG.cell(flexcpChecked, j, Frm.FG.ColIndex("Select")) = flexChecked Then
                            LngAttachedItemID = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemID")))
                            DblAttachedQty = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemQty")))
                        If SystemOptions.Itemsattachedzero = False Then
                            DblAttachedPrice = val(Frm.FG.TextMatrix(j, Frm.FG.ColIndex("AttachItemPrice")))
                         Else
                             DblAttachedPrice = 0
                         End If
                         Grid.TextMatrix(LngRow, Grid.ColIndex("Count")) = Frm.Txtqty
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
                    .cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").value
                    longunitid = RsUnitData("UnitID").value
                    
           If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").value
          Else
                   .TextMatrix(LngRow, .ColIndex("UnitID")) = IIf(IsNull(RsUnitData("UnitNamee").value), RsUnitData("UnitName").value, RsUnitData("UnitNamee").value)
          End If
                Else
                    .cell(flexcpData, LngRow, .ColIndex("UnitID")) = DefalutUnitID
               '     .TextMatrix(LngRow, .ColIndex("UnitID")) = DefalutUnitName
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
            


        If subcall = False Then
            Exit Function
        End If
        
        
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
      
If Grid.ColIndex("RequestTypeName") = -1 Then
    Exit Function
Else
    GetBalnceOfItems val(Me.StoreName.BoundText), LngItemID, LngRow
    
End If
      
   '*************************************
End Function

Public Sub ShowPricess(Optional LngRow As Long, Optional haveoffer As Boolean, Optional discountype As Integer, Optional LngItemID As Long)
Dim discId As Double
Dim discValue As Double
Dim DblQty As Double
Dim TempNo As Double
Dim SngItemPrice As Single
 With Me.Grid
                        If (Me.GridTrans = InvoiceTransaction Or Me.GridTrans = salespricelist) And haveoffer = False Then
                              '  If subcall = True Then GoTo xdis
                              Dim freeitemid As Long
                               Dim freeitemUnitid As Long
                              Dim fQty As Long
                                Dim freeitemQty As Double
                                Dim qtyPercentage As Double
                                Dim netqty As Long
                               DblQty = val(.TextMatrix(LngRow, .ColIndex("Count")))
                               SngItemPrice = val(.TextMatrix(LngRow, .ColIndex("Price")))
                                If discId <= 0 And discountype <> -1 Then discId = discountype
                                            If CheckItem(LngItemID, DtpBillDate, discId, discValue, , val(Branch.BoundText), freeitemid, freeitemUnitid, fQty, freeitemQty, qtyPercentage) = True Then
                                            
                                             If LngRow < Me.Grid.rows - 1 Then
                                               If val(.TextMatrix(LngRow + 1, .ColIndex("Code"))) = freeitemid And val(Grid.TextMatrix(LngRow + 1, Grid.ColIndex("DiscountType"))) = 4 Then
                                               .RemoveItem LngRow + 1
                                                End If
                                            End If
                                                If discountype = 4 Then discId = 4
                                                'If discId <= 0 And discountype <> -1 Then discId = discountype
                                                       If .ColIndex("DiscountType") <> -1 Then
                                                                      
                                                                                       '0 value 1 percentage
                                                                                    If discId = 0 Then
                                                                                    If fQty > 0 Then
                                                                                   TempNo = Math.Round(val(.TextMatrix(LngRow, .ColIndex("Count"))) / fQty, 5)
                                                                                  'NetQty = Math.Truncate(val(.TextMatrix(LngRow, .ColIndex("Count"))) / fQty, 1)
                                                                                  netqty = CDbl(left(CStr(TempNo), 1))

                                                                                '  NetQty = Cel(TempNo)
                                                                                    End If
                                                                                    discValue = netqty * discValue
                                                                                    If discValue <= 0 Then
                                                                                    discId = 4
                                                                                    Else
                                                                                    discId = 2
                                                                                    End If
                                                                                    ElseIf discId = 1 Then
                                                                                      If fQty > 0 Then
                                                                                   netqty = Math.Round(val(.TextMatrix(LngRow, .ColIndex("Count"))) / fQty, 0)
                                                                                  
                                                                                    End If
                                                                                    If netqty <= 0 Then
                                                                                    discValue = 0
                                                                                   discId = 4
                                                                                   Else
                                                                                   discId = 3
                                                                                    End If
                                                                                    
                                                                                       ElseIf discId = 3 Or discId = 2 Then   'over Qty
                                                                                    discId = -2
                                                                                    
                                                                                                 ElseIf discId = 4 Then
                                                                                    discId = 4
                                                                                    
                                                                                    End If
                                                                            If discId <> 4 Then
                                                                              .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = discId
                                                                            .TextMatrix(LngRow, .ColIndex("DiscountType")) = discId
                                                                            .TextMatrix(LngRow, .ColIndex("DiscountVal")) = discValue
                                                                            End If
                                                                                    
                                                                                                                    
                                                                      If discId = -2 Then
                                                                                            .cell(flexcpData, LngRow, .ColIndex("DiscountType")) = 1
                                                                                                .TextMatrix(LngRow, .ColIndex("DiscountType")) = 1
                                                                                                 .TextMatrix(LngRow, .ColIndex("DiscountVal")) = 0
                                                                                                 
                                                                                                   haveoffer = True
                                                                                             If qtyPercentage > 0 Then 'ÿ«·⁄—÷ „‰ ’‰› ⁄·ÌÂ ’‰› «Œ—'
                                                                                           'freeitemQty = (qtyPercentage * DblQty)
                                                                                            
                                                                                            
Dim xx As Double
Dim xx2 As Double
Dim yy As Double
Dim isFree As Boolean
xx = val(DblQty)
xx2 = val(freeitemQty)
 If xx2 <> 0 Then
    yy = xx / xx2
    If freeitemQty <= val(DblQty) Then
        freeitemQty = (qtyPercentage * DblQty)
        freeitemQty = Int(DblQty / fQty) * xx2
        'fQty
    Else
    freeitemQty = freeitemQty * IIf(Int(yy) < 1, 1, Int(yy))
    
    End If
    If Int(yy) <> yy Then
        isFree = True
        
    Else
    isFree = True
    End If
End If


                                                                                                                         If DblQty >= fQty And isFree Then
                                                                                                                        
                                                                                                                      If freeitemQty >= 1 Then
                                                                                                                        AddItemInGrid 0, freeitemid, Int(freeitemQty), SngItemPrice, , , , , , , , , , , , , , 4, , , , haveoffer
                                                                                                                        End If
                                                                                                                    End If
                                                                                             End If
                                                
                                                                          End If
                                          
                                                                        
                                End If
                                            
                   End If
                        
                        
                     
                                
                                
                                
                        End If
 End With
End Sub
Private Sub GetBalnceOfItems(ByVal mStoreNo As Long, ByVal mItemId As Long, ByVal LngRow As Long, Optional ByVal NotChecked As Boolean = False)

If Not NotChecked And val(Grid.TextMatrix(LngRow, Grid.ColIndex("ItemBalance"))) >= val(Grid.TextMatrix(LngRow, Grid.ColIndex("Count"))) Then
    Exit Sub
ElseIf val(Grid.TextMatrix(LngRow, Grid.ColIndex("ItemBalance"))) >= val(Grid.TextMatrix(LngRow, Grid.ColIndex("Count"))) Then
    Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeName")) = ""
    Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeNo")) = ""
    Grid.TextMatrix(LngRow, Grid.ColIndex("StoreIDAvi")) = ""
    Grid.TextMatrix(LngRow, Grid.ColIndex("StoreIDAviName")) = ""
    Exit Sub
End If

Dim StrSQL  As String
Dim mDateTrans As Date
Dim Begin  As Boolean
mDateTrans = XPDtbBill.value



  StrSQL = ""
  
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "        T.StoreID,"
StrSQL = StrSQL & "        ts.StoreName,"
StrSQL = StrSQL & "        ts.storenamee,SUM(T.transqty) as ItemBalance2"
StrSQL = StrSQL & " FROM   ("
StrSQL = StrSQL & "            SELECT SUM("
StrSQL = StrSQL & "                       ("
StrSQL = StrSQL & "                           dbo.Transaction_Details.Quantity / ISNULL("
StrSQL = StrSQL & "                               dbo.GetItemUnitFactor(dbo.Transaction_Details.Item_ID, 1),"
StrSQL = StrSQL & " 1"
StrSQL = StrSQL & "                           ) * StockEffect"
StrSQL = StrSQL & "                       )"
StrSQL = StrSQL & "                   ) AS transqty,"
StrSQL = StrSQL & "                   StockEffect,"
StrSQL = StrSQL & "                   transactions.StoreId"
StrSQL = StrSQL & "            From dbo.transactions"
StrSQL = StrSQL & "                   INNER JOIN dbo.TransactionTypes"
StrSQL = StrSQL & "                        ON  dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
StrSQL = StrSQL & "                   INNER JOIN dbo.Transaction_Details"
StrSQL = StrSQL & "                        ON  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
StrSQL = StrSQL & "            WHERE  dbo.Transactions.Transaction_Date <= " & SQLDate(mDateTrans, True) & ""
StrSQL = StrSQL & "                       Group By"
StrSQL = StrSQL & "                              dbo.Transactions.Transaction_Type,"
StrSQL = StrSQL & "                              dbo.TransactionTypes.StockEffect,"
StrSQL = StrSQL & "                              dbo.Transaction_Details.Item_ID,"
StrSQL = StrSQL & "                              transactions.StoreId"
StrSQL = StrSQL & "                       Having (dbo.TransactionTypes.StockEffect <> 0)"
StrSQL = StrSQL & "                       AND (dbo.Transaction_Details.Item_ID = " & mItemId & " )"
'               --TblItems.ItemID)
'AND dbo.Transactions.StoreID = 1
StrSQL = StrSQL & "                   ) T"
StrSQL = StrSQL & "                   INNER JOIN TblStore AS ts"
StrSQL = StrSQL & "                        ON  ts.StoreID = T.StoreID"
StrSQL = StrSQL & "            Group By"
StrSQL = StrSQL & "                   t.StoreID,"
StrSQL = StrSQL & "                   ts.StoreName,"
StrSQL = StrSQL & "                   ts.storenamee"
StrSQL = StrSQL & "            Having (SUM(t.transqty) >=  " & val(TxtQuantity) & " )"
StrSQL = StrSQL & "            ORDER BY  SUM(T.transqty) DEsc "
Dim rsDummy As New ADODB.Recordset
rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
        Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeName")) = "ÿ·» ‘—«¡ œ«Œ·Ï"
        Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeNo")) = "1"
        Grid.TextMatrix(LngRow, Grid.ColIndex("ItemBalance2")) = ""
        
    Else
        Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeName")) = "ÿ·»  ÕÊÌ· „Œ“‰Ì"
        Grid.TextMatrix(LngRow, Grid.ColIndex("RequestTypeNo")) = "2"
        Grid.TextMatrix(LngRow, Grid.ColIndex("StoreIDAvi")) = rsDummy!StoreID
        Grid.TextMatrix(LngRow, Grid.ColIndex("StoreIDAviName")) = rsDummy!StoreName
        Grid.TextMatrix(LngRow, Grid.ColIndex("ItemBalance2")) = rsDummy!ItemBalance2
    



    End If

     sComboStore = StrSQL
              
End Sub



Sub GetItemBalanceInStore(Optional LngRow As Long, Optional ColorID As Integer, Optional itemsize As Integer, Optional ClassId As Integer, Optional StrItemSerial As String, Optional LngItemID As Long, Optional TransactionDate As Variant)
   If Me.GridTrans = internalorder Or Me.GridTrans = purchaserequest Or Me.GridTrans = salespricelist Or Me.GridTrans = 7 Or GridTrans = ShowPrice Or Me.GridTrans = MoveItems Or (Me.GridTrans = 0) Then
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

        For IntGrdLoopRows = Me.Grid.FixedRows To Me.Grid.rows - 1
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

    If Me.GridTrans <> InvoiceTransaction Then
        Exit Sub
    End If
 '   Load FrmInvProfit
 '   FrmInvProfit.CostMethod = SystemOptions.SysMainStockCostMethod

    StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "TransactionTypes.TransactionTypeName "
    StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN Transactions ON " & "TransactionTypes.Transaction_Type = Transactions.Transaction_Type "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    'FrmInvProfit.FgItems.Rows = Me.Grid.Rows
    'FrmInvProfit.FgItems.FixedRows = Me.Grid.FixedRows

    For i = Me.Grid.FixedRows To Me.Grid.rows - 1
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemID")) = Me.Grid.Cell(flexcpText, i, Me.Grid.ColIndex("Code"))
    
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemCode")) = Me.Grid.Cell(flexcpTextDisplay, i, Me.Grid.ColIndex("Code"))
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("ItemName")) = Me.Grid.Cell(flexcpTextDisplay, i, Me.Grid.ColIndex("Name"))
    '    FrmInvProfit.FgItems.TextMatrix(i, FrmInvProfit.FgItems.ColIndex("CostPrice")) = Me.Grid.TextMatrix(i, Me.Grid.ColIndex("ItemCostPrice"))
'
        If SystemOptions.SysMainStockCostMethod = LastPurPriceType Then
            LngTransID = val(Me.Grid.TextMatrix(i, Me.Grid.ColIndex("PofTransID")))

            If LngTransID <> 0 Then
                If Not (rs.BOF Or rs.EOF) Then
                    rs.Find "Transaction_ID=" & LngTransID & "", , adSearchForward, 1

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

        If Grid.rows > 1 Then
            LblItemsCount.Caption = Grid.Aggregate(flexSTCount, Grid.FixedRows, Grid.ColIndex("Code"), Grid.rows - 1, Grid.ColIndex("Code"))
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

        Do While LngSearchRow < (.rows - 1)
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
            If LngShowRow <= -1 Then
                Exit Sub
            End If
            If LngShowCol <= -1 Then
                Exit Sub
            End If
            If lNoteRow = LngShowRow And lNoteCol = LngShowCol Then
                'If bSetFocus = False Then Exit Sub
            End If

            ClientToScreen Grid.hWnd, uPoint
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
                ShowWindow .hWnd, SW_SHOWNA
                m_QtyToolTipShown = True
            End With

        Else

            With FrmComment
                .PicContainer(0).Visible = False
                .PicContainer(1).Visible = True
                'LngItemID = Val(Me.Grid.TextMatrix(LngShowRow, Grid.ColIndex("Code")))
'                LngItemID = val(StrItemID)
                StrSQL = "SELECT TblItems.PartNo,TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName,TBLiTEMSUnits.MaxSelingPrice,TBLiTEMSUnits.MinSelingPrice, " & "Groups.GroupName, TblItems.SallingPrice, TblItems.CustomerPrice, TblItems.DealerPrice "
                StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
                StrSQL = StrSQL + " Inner join TBLiTEMSUnits On TBLiTEMSUnits.ItemID = TblItems.ItemId and DefaultUnit = 1"
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
                    'sa .LblData(15).Caption = IIf(IsNull(rs("MinSelingPrice").value), "", rs("MinSelingPrice").value)
                   'sa .LblData(16).Caption = IIf(IsNull(rs("MaxSelingPrice").value), "", rs("MaxSelingPrice").value)
                End If

                If Me.GridTrans = InvoiceTransaction Then
                        If val(Customer.BoundText) = 0 Then Customer.BoundText = 2
                     ItemTransInfo = GetLastItemTrans(CLng(StrItemID), 21, val(val(Me.StoreName.BoundText)), val(Customer.BoundText))
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
            
                ShowWindow .hWnd, SW_SHOWNA
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
            .Buttons.Add , "AddRow2", "", , "Plus"
        
            If SystemOptions.UserInterface = ArabicInterface Then
                .Buttons("AddRow").ToolTipText = "≈÷«›… ”ÿ— ÃœÌœ ›Ï ««·›« Ê—…"
                .Buttons("RemoveRow").ToolTipText = "Õ–› «·”ÿ— «·„Õœœ «Ê «·Õ«·Ï ›Ï «·›« Ê—…"
                .Buttons("ClearAll").ToolTipText = "Õ–› ﬂ· «·”ÿÊ— «·„ÊÃÊœ… ›Ï «·›« Ê—…"
                .Buttons("FillData").ToolTipText = "≈œŒ«· «·√’‰«› ⁄‰ ÿ—Ìﬁ „—»⁄ ÕÊ«—"
                .Buttons("Tree").ToolTipText = "≈œŒ«· «·√’‰«› ⁄‰ ÿ—Ìﬁ ‘Ã—… «·√’‰«›"
                '.Buttons("Options").ToolTipText = "≈ŸÂ«— ŒÌ«—«  «·⁄—÷ ›Ï ÃœÊ· «·√’‰«›"
                .Buttons("Barcode").ToolTipText = "ÿ»«⁄… »«—ﬂÊœ ·√’‰«› «·›« Ê—…"
                .Buttons("GridCols").ToolTipText = "≈ŸÂ«— Ê≈Œ›«¡ «·√⁄„œ… ›Ï ÃœÊ· «·√’‰«›"
                .Buttons("AddRow2").ToolTipText = "«÷«›… ”ÿ— „»«‘—… "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .Buttons("AddRow").ToolTipText = "Add new row in grid"
                .Buttons("RemoveRow").ToolTipText = "Remove the current select grid row"
                .Buttons("ClearAll").ToolTipText = "Remove all rows From the grid"
                .Buttons("FillData").ToolTipText = "Enter items in the grid with use the dialog box"
                '.Buttons("Options").ToolTipText = "Show & Hide Columns"
                .Buttons("Tree").ToolTipText = "Enter items in the grid with use the items tree"
                .Buttons("Barcode").ToolTipText = "Print Barcode Label..."
                .Buttons("GridCols").ToolTipText = "Show And Hide Grid Cols"
                .Buttons("AddRow2").ToolTipText = "AddRow2"
            End If

        End With

    End If

End Sub

Public Function NewEmptyRow() As Long
    Dim i As Long
    'Dim LngNewEmptyRow As Long

    For i = Me.Grid.FixedRows To Me.Grid.rows - 1

        If val(Me.Grid.TextMatrix(i, Grid.ColIndex("Name"))) = 0 Then
            NewEmptyRow = i
            Exit Function
        End If

    Next i

    Me.Grid.rows = Me.Grid.rows + 1
    NewEmptyRow = Me.Grid.rows - 1

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
    FG.rows = FG.FixedRows

    Set rs = New ADODB.Recordset
    rs.Open "TblItems", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If Not rs.BOF Or rs.EOF Then
        rs.MoveFirst

        For i = 1 To Me.Grid.rows - 1
            LngItemID = val(Me.Grid.TextMatrix(i, Me.Grid.ColIndex("Name")))

            If LngItemID <> 0 Then
                LngFindRow = FG.FindRow(LngItemID, FG.FixedRows, FG.ColIndex("ItemID"), False, True)

                If LngFindRow = -1 Then
                    rs.MoveFirst
                    rs.Find "ItemID=" & LngItemID & "", , adSearchForward, 1

                    If Not (rs.BOF Or rs.EOF) Then

                        With FG
                            .rows = .rows + 1
                            .TextMatrix(.rows - 1, .ColIndex("Serial")) = .rows - 1
                            .TextMatrix(.rows - 1, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                            .TextMatrix(.rows - 1, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                            .TextMatrix(.rows - 1, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                            .TextMatrix(.rows - 1, .ColIndex("SallingPrice")) = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                            .TextMatrix(.rows - 1, .ColIndex("CustomerPrice")) = IIf(IsNull(rs("CustomerPrice").value), "", rs("CustomerPrice").value)
                            .TextMatrix(.rows - 1, .ColIndex("DealerPrice")) = IIf(IsNull(rs("DealerPrice").value), "", rs("DealerPrice").value)
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
    
    Dim mStoreNameString As String
    mStoreNameString = Me.StoreName.Text
    
    CheckDataEntered = False
    
   If Not GridQuantity Is Nothing Then
    
    GridQuantity.rows = 1
GridQuantity.rows = Grid.rows
Dim isFound As Boolean
Dim j As Long, i As Long




With GridQuantity
    For i = 1 To Grid.rows - 1
        
        For j = 1 To .rows - 1
            If val(.TextMatrix(j, .ColIndex("ItemID"))) = val(Grid.TextMatrix(i, Grid.ColIndex("Code"))) And j <> i Then
                .TextMatrix(j, .ColIndex("Quantity")) = val(.TextMatrix(j, .ColIndex("Quantity"))) + val(Grid.TextMatrix(i, Grid.ColIndex("Count")))
                isFound = True
                IsDuplicateItem = True
                GoTo NextRow
            Else
            
            End If
        Next j
        
NextRow:
            If Not isFound Then
                .TextMatrix(i, .ColIndex("ItemID")) = val(Grid.TextMatrix(i, Grid.ColIndex("Code")))
                .TextMatrix(i, .ColIndex("Name")) = Grid.ValueMatrix(i, Grid.ColIndex("Name"))
                .TextMatrix(i, .ColIndex("UnitId")) = Grid.cell(flexcpData, i, Grid.ColIndex("UnitID"))
                .TextMatrix(i, .ColIndex("UnitName")) = Grid.TextMatrix(i, Grid.ColIndex("UnitId"))
                
                .TextMatrix(i, .ColIndex("ColorID")) = Grid.TextMatrix(i, Grid.ColIndex("ColorID"))
                .TextMatrix(i, .ColIndex("ItemSize")) = Grid.TextMatrix(i, Grid.ColIndex("ItemSize"))
                .TextMatrix(i, .ColIndex("ClassId")) = Grid.TextMatrix(i, Grid.ColIndex("ClassId"))
                .TextMatrix(i, .ColIndex("Serial")) = Grid.TextMatrix(i, Grid.ColIndex("Serial"))
                .TextMatrix(i, .ColIndex("ItemType")) = Grid.TextMatrix(i, Grid.ColIndex("ItemType"))
                .TextMatrix(i, .ColIndex("Name")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Name"))
                .TextMatrix(i, .ColIndex("Code")) = Grid.cell(flexcpTextDisplay, i, Grid.ColIndex("Code"))
                
                .TextMatrix(i, .ColIndex("Quantity")) = Grid.TextMatrix(i, Grid.ColIndex("Count"))
            End If
            isFound = False
    Next
End With
    
End If
'Debug.Print val(Me.storename.BoundText)
'GheckLinkItem
    '«·ﬂ‘› ⁄·Ï «‰ «·›« Ê—…  Õ ÊÏ ⁄·Ï √’‰«›
'    If Me.Grid.FixedRows = Me.Grid.rows Or ModFgLib.ItemsInGrid(Me.Grid, Me.Grid.ColIndex("Code")) = -1 Then
'  If SystemOptions.UserInterface = ArabicInterface Then
'        Msg = "·« ÊÃœ √’‰«› ›Ï  Â–« «·”‰œ ...!!!"
'   Else
'   Msg = "No Items In Voucher ...!!!"
'
'   End If
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If



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

        For LngRowNum = Me.Grid.FixedRows To Me.Grid.rows - 1

            If val(.TextMatrix(LngRowNum, .ColIndex("Code"))) <> 0 Then
                StrCurrentItemName = Grid.cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
  LngUnitID = val(Grid.cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")))
                If LngUnitID = 0 Then
                 GetDefaultItemUnit LngItemID, LngUnitID
                 Grid.TextMatrix(LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 Grid.cell(flexcpData, LngRowNum, Grid.ColIndex("UnitID")) = LngUnitID
                 
                End If
'If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType"))) = "" Then '·«»œ Ê«‰ Ìﬂ..Ê‰ «·’‰› ”·⁄…
   Dim rsDummyIte As New ADODB.Recordset
   Dim s As String
   s = "Select ItemType from tblItems where ItemId = " & LngItemID
   Set rsDummyIte = New ADODB.Recordset
   rsDummyIte.Open s, Cn, adOpenForwardOnly, adLockReadOnly
   If Not rsDummyIte.EOF Then
        Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType")) = val(rsDummyIte!ItemType & "")
   End If
'End If
'---------------«· «ﬂœ „‰ «‰ „œ… «·⁄—÷ ﬁÌ„… —ﬁ„Ì…
If Grid.ColIndex("PODaysSub") <> -1 Then
             If Not IsNumeric(val((Grid.TextMatrix(LngRowNum, Grid.ColIndex("PODaysSub"))))) Then
                                      Msg = " «—ÌŒ «· ”·Ì„  " & CHR(13)
                                      Msg = Msg + StrCurrentItemName & CHR(13)
                                      Msg = Msg + "€Ì— —ﬁ„Ì…"
                                      MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                      FocusGridOnCell LngRowNum, Grid.ColIndex("PODaysSub")
                                      Screen.MousePointer = vbDefault
                                      Exit Function
            End If
End If
                '----------«· «ﬂœ „‰ ≈œŒ«· ﬂ„Ì… «·’‰›
                If val(.TextMatrix(LngRowNum, .ColIndex("Count"))) <= 0 And Me.GridTrans <> NewGard Then
'                    Msg = "„‰ ›÷·ﬂ ÌÃ»  ÕœÌœ «·ﬂ„Ì… «·Œ«’… »«·’‰›" & CHR(13)
'                    Msg = Msg + StrCurrentItemName & CHR(13)
'                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                    FocusGridOnCell LngRowNum, Grid.ColIndex("Count")
'                    Screen.MousePointer = vbDefault
'                    Exit Function
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
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Grid.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If

                    '-------«· «ﬂœ „‰ ≈œŒ«· ‰Ê⁄ «·Œ’„ ⁄·Ï «·»—‰«„Ã
                ElseIf Grid.ColIndex("DiscountType") <> -1 Then
                    If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = "" Then
                        Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType")) = 0
                    End If
                    If (Grid.TextMatrix(LngRowNum, Grid.ColIndex("DiscountType"))) = "" Then
                        
                                If SystemOptions.UserInterface = ArabicInterface Then
                          Msg = "„‰ ›÷·ﬂ Õœœ ‰Ê⁄ «·Œ’„ «·Œ«’ »«·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Discount Type For Item" & CHR(13)
                        End If
                        
                        
                        Msg = Msg + StrCurrentItemName & CHR(13)
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
                           Msg = "„‰ ›÷·ﬂ Õœœ ﬁÌ„… «·Œ’„ «·Œ«’ »«·’‰›" & CHR(13)
                        Else
                        Msg = "Must Enter Discount Value For Item" & CHR(13)
                        End If
                        
                            Msg = Msg + StrCurrentItemName & CHR(13)
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
                       
                       
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            FocusGridOnCell LngRowNum, Grid.ColIndex("DiscountVal")
                            Screen.MousePointer = vbDefault
                            Exit Function
                        End If
                    End If
                End If

                If Me.GridTrans <> ShowPrice Then

                    '-------„—«Ã⁄… »Ì«‰«  «·”Ì—Ì«· «·„œŒ·
                    If .cell(flexcpChecked, LngRowNum, .ColIndex("HaveSerial")) = flexChecked Then

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
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
                            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        
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
                                If Not (IsNumeric(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) < 1 Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) > 36 Or val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime"))) <> Int(val(.TextMatrix(LngRowNum, .ColIndex("guaranteeTime")))) Then
                                    Msg = "„œ… «·÷„«‰ «·Œ«’… »«·’‰›" & CHR(13)
                                    Msg = Msg + StrCurrentItemName & CHR(13)
                                    Msg = Msg + "€Ì— ’«·Õ…" & CHR(13)
                                    Msg = Msg + "√œŒ· „œ…   —«ÊÕ »Ì‰ 1 : 36 ‘Â—" & CHR(13)
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
        For LngRowNum = 1 To Grid.rows - 1

            If Trim$(Me.Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) <> "" Then
                LngSearchResault = Grid.FindRow(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")), , 4)

                If LngSearchResault <> LngRowNum And LngSearchResault > 0 Then
                    StrTemp = Grid.cell(flexcpTextDisplay, LngSearchResault, Grid.ColIndex("Name"))
                    Msg = " „ «œŒ«· «·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                    Msg = Msg + StrCurrentItemName & CHR(13)
                    Msg = Msg + "›Ï «·”ÿ—:" & LngRowNum
                    Msg = Msg + "ﬁ»· –·ﬂ «·Ì «·’‰›" & CHR(13)
                    Msg = Msg + StrTemp & CHR(13)
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

            For LngRowNum = 1 To Grid.rows - 1

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
                                    Msg = Msg & "[" & mStoreNameString & "]"
                                End If

                                Msg = Msg & CHR(13) & "------------------------------------"
                                Msg = Msg & CHR(13) & "Õ«Ê· «‰  ” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ ”Ì—Ì«· ·„⁄—›… «·„“Ìœ"
                                Msg = Msg & CHR(13) & "Â·  —Ìœ «‰ ÌﬁÊ„ «·»—‰«„Ã »«·»ÕÀ ·ﬂ ⁄‰  Â–« «·”Ì—Ì«· «·¬‰...!!!"
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

'delme
'SelectedIssueVoucher = False
        If SelectedIssueVoucher = True Then
            CheckDataEntered = True
            Exit Function
        End If

        If SystemOptions.TypicalProduction = False Then
            'If (Me.GridTrans = 2 Or Me.GridTrans = 14 Or Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = ReturnTransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = Destruction) Then
If (Me.GridTrans = 2 Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction) Then
                For LngRowNum = 1 To Grid.rows - 1
                    CalculateInvProfit LngRowNum
                Next

            End If
 
        End If
        
        If (Me.GridTrans = 31 Or Me.GridTrans = 2 Or Me.GridTrans = MoveItems Or Me.GridTrans = 14 Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Destruction Or (Me.GridTrans = GridTransType.internalorder And SystemOptions.NotAllowStockNegativeInternal)) Then
            If (Me.GridTrans = 31 Or Me.GridTrans = 2 Or Me.GridTrans = 14 Or Me.GridTrans = InventoryOut Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = InvoiceTransaction Or Me.GridTrans = Returntransaction Or Me.GridTrans = MoveItems Or Me.GridTrans = Destruction Or (Me.GridTrans = GridTransType.internalorder And SystemOptions.NotAllowStockNegativeInternal)) Then
                If SystemOptions.SysAllowStockNegative = False Then

                    For LngRowNum = 1 To Grid.rows - 1
                
                        If val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemType"))) = 0 Then '·«»œ Ê«‰ Ìﬂ..Ê‰ «·’‰› ”·⁄…
                            If Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")) <> "" Then
                                '-------------------------------------------------------------------
                                '«·»Ì«‰«  «·√”«”Ì… —ﬁ„ «·’‰› Ê «”„ «·’‰› Ê «·”Ì—Ì«·
                                LngItemID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Code")))
                               ' LngItemID = val(GridQuantity.TextMatrix(LngRowNum, GridQuantity.ColIndex("ItemID")))
                                StrCurrentItemName = Grid.cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Name"))
                               ' StrCurrentItemName = val(GridQuantity.TextMatrix(LngRowNum, GridQuantity.ColIndex("ItemID")))
                                StrCurrentItemCode = Grid.cell(flexcpTextDisplay, LngRowNum, Grid.ColIndex("Code"))
                               ' StrCurrentItemCode = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemID")))
                                LngColorID = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ColorID")))
                                StrItemSize = Grid.TextMatrix(LngRowNum, Grid.ColIndex("ItemSize"))
                                LngClassId = val(Grid.TextMatrix(LngRowNum, Grid.ColIndex("ClassId")))

                                If Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial"))) <> "" Then
                                    StrItemSerial = Trim(Grid.TextMatrix(LngRowNum, Grid.ColIndex("Serial")))
                                Else
                                    StrItemSerial = ""
                                End If

                                
                                DblOutQty = CheckTotalQty(LngRowNum, LngColorID, StrItemSize, LngClassId, StrItemSerial, LngRowNum)
                       
                                '-------------------------------------------------------------------
                                If LngColorID = 0 Then LngColorID = 1
                                 If LngClassId = 0 Then LngClassId = 1
                                 If StrItemSize = "" Then StrItemSize = "1"
                                 
                                 
                                If Me.TxtModFlag.Text = "N" Then
                                    If (Me.GridTrans = GridTransType.internalorder And SystemOptions.NotAllowStockNegativeInternal) Then
                                        Set RsTest = GetItemQuantityStock(LngItemID, val(val(DefStorePurchase)), Me.DtpBillDate.value, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                            
                                            mStoreNameString = mStorePurName
                                    Else
                                        Set RsTest = GetItemQuantityStock(LngItemID, val(val(Me.StoreName.BoundText)), Me.DtpBillDate.value, , , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                    End If
                                ElseIf Me.TxtModFlag.Text = "E" Then
                                    '·«ÕŸ «‰Â ›Ï Õ«·… «· ⁄œÌ· ·«»œ „‰ ≈—”«· —ﬁ„ «·›« Ê—… ··œ«·…
                        
                                    If (Me.GridTrans = InvoiceTransaction) Or (Me.GridTrans = Returntransaction) Then
                      
                                        Set RsTest = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), Me.DtpBillDate.value, val(Me.TxtNots.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                        
                                        'Set RsTest = GetItemQuantityStock(LngItemID, val(Me.storename.BoundText), _
                                        'Me.DtpBillDate.value,  Val(Me.TxtInvID.text), , , StrItemSerial, True, LngColorID, StrItemSize)
                                    Else
                                        If (Me.GridTrans = GridTransType.internalorder And SystemOptions.NotAllowStockNegativeInternal) Then
                                            Set RsTest = GetItemQuantityStock(LngItemID, val(DefStorePurchase), Me.DtpBillDate.value, val(Me.TxtInvID.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                            mStoreNameString = mStorePurName
                                            
                                        Else
                                            Set RsTest = GetItemQuantityStock(LngItemID, val(Me.StoreName.BoundText), Me.DtpBillDate.value, val(Me.TxtInvID.Text), , , StrItemSerial, True, LngColorID, StrItemSize, LngClassId)
                                        End If
                        
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
                                    Msg = Msg & "·« ÊÃœ „‰Â «Ï ﬂ„Ì… ›Ï «·„Œ“‰ : " & mStoreNameString & CHR(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & CHR(13)
                                    Msg = Msg & "--------------------------" & CHR(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & CHR(13) & ""
                                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                    FocusGridOnCell LngRowNum, .ColIndex("Name")
                                    TerminateRecordset RsTest
                                    Exit Function
                                ElseIf (Round(RsTest("totalqty").value, 2) < Round(DblOutQty, 2)) And StrItemSerial = "" Then
                                    Msg = "«·’‰› : " & StrCurrentItemName & CHR(13)
                                    Msg = Msg & "ﬂÊœ —ﬁ„ : " & StrCurrentItemCode & CHR(13)
                                    Msg = Msg & "«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & LngRowNum & CHR(13)
                                    Msg = Msg & "≈Ã„«·Ï «·ﬂ„Ì… «·„ÊÃÊœ… „‰Â ›Ï «·„Œ“‰ : " & mStoreNameString & CHR(13)
                                    Msg = Msg & "Ê–·ﬂ Õ Ï  «—ÌŒ : " & DisplayDate(DtpBillDate.value) & CHR(13)
                                    Msg = Msg & "ÂÏ : " & RsTest("totalqty").value & CHR(13)
                                    Msg = Msg & "«·ﬂ„Ì… «·„—«œ ’—›Â« : " & DblOutQty & CHR(13)
                                    Msg = Msg & "·«Ì„ﬂ‰ «·”„«Õ »Œ—ÊÃ Â–Â «·ﬂ„Ì….  " & CHR(13)
                                    Msg = Msg & "--------------------------" & CHR(13) & ""
                                    Msg = Msg & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…" & CHR(13) & ""
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




Private Function CheckTotalQty(LngRow As Long, _
                               Optional LngColorID1 As Long, _
                               Optional StrItemSize1 As String, _
                               Optional LngClassId1 As Long, _
                               Optional StrItemSerial1 As String, Optional ByVal currentrow As Long) As Double
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
Dim StrIndex As Long
    If LngRow = 0 Then
        Exit Function
    End If

    With Me.Grid
        LngFoundGridRow = 0
        
        
If currentrow = 0 Then
currentrow = .rows - 1
StrIndex = 1
Else
StrIndex = currentrow
End If

If IsDuplicateItem Then
    currentrow = .rows - 1
    StrIndex = 1
End If
        For i = StrIndex To currentrow
            LngFoundGridRow = .FindRow(.TextMatrix(LngRow, .ColIndex("Code")), i, .ColIndex("Code"), True, True)

            If LngFoundGridRow > 0 Then
                i = LngFoundGridRow
            
                LngCurItemID = .TextMatrix(LngFoundGridRow, .ColIndex("Code"))
                LngUnitID = val(.cell(flexcpData, LngFoundGridRow, .ColIndex("UnitID")))
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
        If UCase(frmname) = "FRMSALEBILL6" Then
            Frm.RetrunType = 66
        End If
        Frm.LngRow = Row
    
        'Frm.LngCol = Col
    
        '' Frm.XPDtbBill.value = Date
 
    End With

    Frm.show vbModal
End Sub
  
Private Sub ItemLotNoFun(Row As Long, _
                         Col As Long, _
                         Item_ID As Long, _
                         StoreID As Long, _
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
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
        My_SQL = My_SQL & " Where (dbo.TransactionTypes.StockEffect <> 0) and (dbo.Transactions.StoreID = " & StoreID & ") And (dbo.Transaction_Details.Item_ID = " & Item_ID & ")"
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
                .rows = 2
                .Clear flexClearScrollable

                If rs.RecordCount > 0 Then
                    .rows = rs.RecordCount + 1
                    rs.MoveFirst

                    For i = 1 To .rows - 1
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
            Frm.cmdOK.Enabled = False
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Item_ID = Me.Grid.TextMatrix(Row, .ColIndex("Code"))

 

 
    
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
    
        Frm.GranteeStartDate.value = Date
    
        If 1 = 1 Then
            If Me.TxtModFlag.Text = "R" Then
                Frm.TxtComment.locked = True
                Frm.cmdOK.Enabled = False
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Item_ID = val(Me.Grid.TextMatrix(Row, .ColIndex("Code")))

 Frm.TxtUnitID = .cell(flexcpData, .Row, .ColIndex("UnitID"))
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
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
            Frm.cmdOK.Enabled = False
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Frm.ItemID.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Name"))
        
        Frm.Vatyo.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Vatyo"))
        Frm.Vat.Caption = Me.Grid.TextMatrix(Row, .ColIndex("Name"))
        
                 Frm.lblunitid.Caption = .cell(flexcpData, .Row, .ColIndex("UnitID"))
            Frm.lblunitname.Caption = .TextMatrix(.Row, .ColIndex("UnitId"))
            
Frm.TxtPrice.Text = .TextMatrix(.Row, .ColIndex("Price"))


        Set Frm.FG = Me.Grid
    
        Frm.LngRow = Row
        Frm.LngCol = Col
        '   Frm.TxtComment.text = Trim$(Me.Grid.TextMatrix(Row, Col))
        Frm.TxtComment.Text = ""
   
        Dim i As Integer

        For i = 1 To Grid.rows - 1

            If Me.Grid.cell(flexcpTextDisplay, i, .ColIndex("Code")) = Frm.lbl(4).Caption Then
                If i < Grid.rows - 1 Then
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
            Frm.cmdOK.Enabled = False
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
        Frm.lbl(4).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Code"))
        Frm.lbl(5).Caption = Me.Grid.cell(flexcpTextDisplay, Row, .ColIndex("Name"))
        Set Frm.FG = Me.Grid
        Frm.LngRow = Row
        Frm.LngCol = Col
        Frm.TxtComment.Text = Trim$(Me.Grid.TextMatrix(Row, Col))

        If Me.TxtModFlag.Text = "R" Then
            Frm.TxtComment.locked = True
            Frm.cmdOK.Enabled = False
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

Private Sub txtCallingID_Change()

    If Me.TxtModFlag.Text = "E" Or Me.TxtModFlag.Text = "N" Then
           Dim s As String
    Dim Num As Long
    Dim mPrice  As Double
    If val(txtCallingID) <> 0 Then
       Dim rsDummy As New ADODB.Recordset
       
        s = " SELECT "
        s = s & "        TblStudCalling.ID        AS ReservNo,"
        s = s & "               TblStudCalling.EnterDate,"
        s = s & "               HoursT                   AS TIMER,"
        s = s & "               HoursT                   AS Hours,"
        s = s & "               TblCustemers.CusName,"
        s = s & "               TblCustemers.CusID,"
        s = s & "               TblEmployee.Emp_Name        EmpName,TblEmployee.Emp_ID,"
        s = s & "               TblItems.ItemName,TblItems.ItemID,tblitems.Fullcode,"
        s = s & "               tblReservationType.Name  AS ReservationTypeName"
        s = s & "        From TblStudCalling2"
        s = s & "               Left outer join  tblReservationType"
        s = s & "                    ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
        s = s & "               INNER JOIN TblEmployee"
        s = s & "                    ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
        s = s & "               INNER JOIN TblStudCalling"
        s = s & "                    ON  TblStudCalling.ID = TblStudCalling2.MasterID"
        s = s & "               INNER JOIN TblCustemers"
        s = s & "                    ON  TblCustemers.CusID = TblStudCalling.CompID"
        s = s & "               INNER JOIN TblItems"
        s = s & "                    ON  TblItems.ItemID = TblStudCalling2.ItemID"
        s = s & "              WHERE  TblStudCalling.ID = " & val(txtCallingID)
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
         If Not (rsDummy.EOF Or rsDummy.BOF) Then
            'TxtCashCustomerName = rsDummy!CusName & ""
            Customer.BoundText = val(rsDummy!CusID & "")
            Grid.rows = rsDummy.RecordCount + 1
    
            For Num = 1 To rsDummy.RecordCount
                AddItemInGrid Num, IIf(IsNull(rsDummy("ItemID")), "", (rsDummy("ItemID").value)), 1, , , , , 1, , , , , , , , , , , val(rsDummy!Emp_id & "")
                rsDummy.MoveNext
                
            Next
            
        End If
    End If    'If Not lvwMain.SelectedItem Is Nothing Then

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
Dim mLblValueAdded As Double
Dim mTaxExemptTotal As Double
mTaxExemptTotal = 0
ItemCount = 0

'LblValueAdded.Tag = 0
    With Me.Grid

        If m_ItemType = ItemsAllType Then

            For RowNum = .FixedRows To .rows - 1
            ItemCount = ItemCount + 1
            linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
                    If .ValueMatrix(RowNum, .ColIndex("chkTaxExempt")) = True Then
                        'mTaxExemptTotal = mTaxExemptTotal + ((val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price")))))
                        mTaxExemptTotal = mTaxExemptTotal + ((val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))) - val(.TextMatrix(RowNum, .ColIndex("DiscountValue")))
                     End If
            
            Next RowNum

        ElseIf m_ItemType = ItemsGoodType Then

            For RowNum = .FixedRows To .rows - 1

                If val(.TextMatrix(RowNum, .ColIndex("ItemType"))) = 0 And val(.TextMatrix(RowNum, .ColIndex("Name"))) <> 0 Then
                ItemCount = ItemCount + 1
                     linetotl = (val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))
                     If .ValueMatrix(RowNum, .ColIndex("chkTaxExempt")) = True Then
                        mTaxExemptTotal = mTaxExemptTotal + ((val(.TextMatrix(RowNum, .ColIndex("Count"))) * val(.TextMatrix(RowNum, .ColIndex("Price"))))) - val(.TextMatrix(RowNum, .ColIndex("DiscountValue")))
                     End If
                     If SystemOptions.PursgaseWithoutDecimal = True And Me.GridTrans = PurchaseTransaction Then
                         DblTemp = DblTemp + Int(linetotl)
                     Else
                           DblTemp = DblTemp + Round(linetotl, SystemOptions.SysDefCurrencyForamt)
                    End If
                    
                End If

            Next RowNum

        ElseIf m_ItemType = ItemsServiceType Then

            For RowNum = .FixedRows To .rows - 1

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

            For RowNum = .FixedRows To .rows - 1
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
'    LblValueAdded.Tag = mTaxExemptTotal
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

        For RowNum = .FixedRows To .rows - 1
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

        For i = 1 To .rows - 1
            SngOutQty = 0
            SngMaxOutQty = 0
            SngWillGoOutQty = 0
        
            LngSearchItemID = val(.TextMatrix(i, .ColIndex("Name")))
            SngWillGoOutQty = val(.TextMatrix(i, .ColIndex("Count")))

            If BolThereIsOutBefore = True Then
                RsOutQty.MoveFirst
                RsOutQty.Find "ItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

                If Not (RsOutQty.BOF Or RsOutQty.EOF) Then
                    SngOutQty = IIf(IsNull(RsOutQty("SumQty").value), 0, RsOutQty("SumQty").value)
                Else
                    SngOutQty = 0
                End If

            Else
                SngOutQty = 0
            End If

            rs.MoveFirst
            rs.Find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

            If Not (rs.BOF Or rs.EOF) Then
                SngMaxOutQty = rs("PartItemQty").value

                If (SngWillGoOutQty + SngOutQty) > SngMaxOutQty Then
                    Msg = "«·’‰›..." & rs("ItemName").value
                    Msg = Msg & CHR(13) & "√ﬁ’Ì ﬂ„Ì… „‰Â ›Ï «·’‰› «·„Ã„⁄ :" & SngMaxOutQty

                    If IntAllowMaxOverOut = -1 Then
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…: ·„ Ìﬁ„ „œÌ— «·‰Ÿ«„ »≈⁄œ«œ ’·«ÕÌ« ﬂ ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                        Msg = Msg & CHR(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        CheckRestrictedItemQty = False
                        Exit Function
                    ElseIf IntAllowMaxOverOut = 0 Then
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…: ·Ì” ·ﬂ ’·«ÕÌ… ·’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄ ›Ï «„— «·‘€·"
                        Msg = Msg & CHR(13) & "·„ Ì”„Õ ·ﬂ » ”ÃÌ· «„— «·‘€·"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        CheckRestrictedItemQty = False
                        Exit Function
                    ElseIf IntAllowMaxOverOut = 1 Then
                        Msg = Msg & CHR(13) & "Â·  —Ìœ ’—› ﬂ„Ì«  «ﬂ»— „‰ «·’‰› «·„Ã„⁄"
                        IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CheckRestrictedItem = False
        Exit Function
    End If

    With Me.Grid

        For i = .FixedRows To .rows - 1
            LngSearchItemID = val(.TextMatrix(i, .ColIndex("Name")))

            If LngSearchItemID <> 0 Then
                rs.MoveFirst
                rs.Find "PartItemID=" & LngSearchItemID, , adSearchForward, adBookmarkFirst

                If rs.BOF Or rs.EOF Then
                    StrTemp = "«·’‰› ..." & .cell(flexcpTextDisplay, i, .ColIndex("Name"))
                    Msg = StrTemp
                    Msg = Msg & CHR(13) & "€Ì— „ÊÃÊœ ›Ï «·’‰› «·„Ã„⁄ «·„— »ÿ »Â «„— «·‘€·"
                    Msg = Msg & CHR(13) & ""
                    Msg = Msg & CHR(13) & "Â·  —Ìœ «·„ «»⁄… ›Ï Õ›Ÿ «„— «·‘€·...øøø"
                    IntRes = MsgBox(Msg, vbExclamation + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

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
        LngUnitID = val(.cell(flexcpData, LngRow, .ColIndex("UnitID")))
    
        If LngItemID = 0 Then
            Exit Sub
        End If
   
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
 Public Sub txtManulaVat_Change()
 
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
              For i = 1 To .rows - 1
    
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


Private Sub XPTxtDiscountVal_Change()
SentTypeVAT
End Sub

Private Sub XPTxtDiscountVal_KeyDown(KeyCode As Integer, Shift As Integer)
If SystemOptions.DiscountByQtyOnly And val(XPTxtDiscountVal.Text) <> 0 Then
    Dim i As Long
    
    For i = 1 To Grid.rows - 1
        Calculate i
    Next
End If
SentTypeVAT
End Sub

Private Sub XPTxtDiscountVal_Validate(Cancel As Boolean)
If SystemOptions.DiscountByQtyOnly And val(XPTxtDiscountVal.Text) <> 0 Then
    Dim i As Long
    
    For i = 1 To Grid.rows - 1
        Calculate i
    Next
End If
SentTypeVAT
End Sub
