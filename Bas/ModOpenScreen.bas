Attribute VB_Name = "ModOpenScreen"

Public Enum ScreensName
    EmployeesScreen '‘«‘… »Ì«‰«  «·„ÊŸ›Ì‰
    CustomersScreen '‘«‘… »Ì«‰«  «·⁄„·«¡
    SuppliersScreen '‘«‘… »Ì«‰«  «·„Ê—œÌ‰
    OtherCustomersScreen '»Ì«‰«  «·„ ⁄«„·Ê‰
    ManCompaniesScreen '»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…
    ItemsGroupsScreen '‘«‘… „Ã„Ê⁄«  «·√’‰«›
    ItemsUnitsScreen '‘«‘… ÊÕœ«  «·√’‰«›
    ItemsDataScreen ' ‘«‘… »Ì«‰«  «·√’‰«›
    StoresDataScreen '‘«‘… »Ì«‰«  «·„Œ«“‰
    BanksDataScreen '‘«‘… »Ì«‰«  «·»‰Êﬂ
    BoxesDataScreen '‘«‘… »Ì«‰«  «·Œ“‰
    WorkOrdersDataScreen '‘«‘… »Ì«‰«  √Ê«„— «·‘€·
    CurrencyDataScreen '‘«‘… »Ì«‰«  «·⁄„·« 
    ShowPriceScreen '⁄—÷ «·”⁄—
    TemplateScreen '«·⁄—Ê÷ «·Ã«Â“…
    InvoiceScreen '‘«‘… «·›« Ê—…
    PurchaseScreen '›« Ê—… «·„‘ —Ì« 
    RetrunSalles '„— Ã⁄ «·„»Ì⁄« 
    RetrunPurchse '„— Ã⁄ «·„‘ —Ì« 
    OpenStockBalance '‘«‘… «·—’Ìœ«·√›  «ÕÏ
    MaintainceGoOnScreen '‘«‘… «·œŒÊ· ··’Ì«‰…
    DestructionScreen '‘«‘… «· ·›Ì« 
    StockCountScreen 'Ã—œ «·„Œ“Ê‰
    StockTransfereScreen ' ÕÊÌ· „‰ „Œ“‰ ·„Œ“‰
    StockSettlementScreen ' ”ÊÌ… «·„Œ“Ê‰
    CheckItemQty '«·√” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›
    CheckItemswaped ' «·«” ⁄·«„ ⁄‰ «·»œ«∆·
    CheckItemSerial
    ExpensesTypes '‘«‘… «‰Ê«⁄ «·„’—Ê›« 
    RevenuesTypes ' ‘«‘… ≈‰Ê«⁄ «·≈Ì—«œ« 
    ExpensesDataScreen '‘«‘… «·„’—Ê›« 
    PaymentsDataScreen '‘«‘… »Ì«‰«  «·„œ›Ê⁄« 
    CashingDataScreen '‘«‘… »Ì«‰«  «·„ﬁ»Ê÷« 
    AllowsDiscountsScreen '‘«‘… «·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…
    ReceiptPartScreen '‘«‘…  Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ
    BoxesStockScreen '‘«‘… Ã—œ «·Œ“‰…
    PopUpShowPaymentTime '⁄—÷ «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…
    PopUpShowInstallmentMustPay '⁄—÷ «·√ﬁ”«ÿ «·„” Õﬁ…
    BarCodeDesign '‘«‘…  ’„Ì„ «·»«—ﬂÊœ
    DayReports ' ﬁ«—Ì— Ê√Õœ«  «·ÌÊ„
    ItemsMainPriceLise 'ﬁ«∆„… «”⁄«— «·√’‰«›
    ItemsPricePlane 'Œÿ…  ”⁄Ì— «·√’‰«›
    OptionsScreen '‘«‘… ŒÌ«—«  «·»—‰«„Ã
    CustomerFile '‘«‘… „·› «·⁄„Ì·
    ReportsManger '‘«‘… „œÌ— «· ﬁ«—Ì—
    StatisticsShow '‘«‘… «·≈Õ’«∆Ì« 
    PopUpShowCusBalances '‘«‘… «·≈” ⁄·«„ ⁄‰ √—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰
    PopUpShowItemsRequest '‘«‘… «·√’‰«› «· Ï »·€  Õœ «·ÿ·»
    PopUpShowItemQuantity '«·≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›
    PopUpShowBoxesAccounts '‘«‘… ⁄—÷ «—’œ… «·Œ“‰
    PopUpShowGuaranteeAlram '‘«‘…  ‰»ÌÂ ÷„«‰ «·√’‰«›
    PopUpSowStagnantItems ' ‰ÌÂ «·—«ﬂœ…
    PopUpShowStockMovement '  ‰»ÌÂ Õ—ﬂ… «·„Œ“Ê‰
    PopUpShowItemCardScreen
    PopUpShowCustomerBalanceScreen
    PopUpShowItemCostScreen '‘«‘… ⁄—÷ „ Ê”ÿ  ﬂ·›… «·’‰›
End Enum

Private Function GetFormName(ScreenName As ScreensName) As String
    Dim StrTempFormName As String

    Select Case ScreenName
        Case ScreensName.CheckItemswaped
'»Ì«‰«  «·«’‰« › «·»œÌ·…
          StrTempFormName = "FrmSearchSerial1"

        Case ScreensName.EmployeesScreen
            '»Ì«‰«  «·„ÊŸ›Ì‰
            StrTempFormName = "FrmEmployee"

        Case ScreensName.CustomersScreen
            '»Ì«‰«  «·⁄„·«¡
            StrTempFormName = "FrmCustemers"

        Case ScreensName.SuppliersScreen
            '»Ì«‰«  «·„Ê—œÌ‰
            StrTempFormName = "FrmCompany"

        Case ScreensName.OtherCustomersScreen
            '»Ì«‰«  «·„ ⁄«„·Ê‰
            StrTempFormName = "FrmOtherCustomers"

        Case ScreensName.ManCompaniesScreen
            '»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…
            StrTempFormName = "FrmManCompanies"

        Case ScreensName.ItemsGroupsScreen
            '»Ì«‰«  „Ã„Ê⁄«  «·√’‰«›
            StrTempFormName = "FrmGroups"

        Case ScreensName.ItemsUnitsScreen
            '»Ì«‰«  ÊÕœ«  «·√’‰«›
            StrTempFormName = "FrmSystemUnites"

        Case ScreensName.ItemsDataScreen
            '»Ì«‰«  «·√’‰«›
            StrTempFormName = "FrmItems"

        Case ScreensName.StoresDataScreen
            '»Ì«‰«  «·„Œ«“‰
            StrTempFormName = "FrmStoreData"

        Case ScreensName.BanksDataScreen
            '»Ì«‰«  «·»‰Êﬂ
            StrTempFormName = "FrmBanksData"

        Case ScreensName.BoxesDataScreen
            '»Ì‰«‰«  «·Œ“‰
            StrTempFormName = "FrmBoxesData"

        Case ScreensName.CurrencyDataScreen
            '»Ì«‰«  «·⁄„·« 
            StrTempFormName = "FrmCurrencyData"

        Case ScreensName.WorkOrdersDataScreen
            '»Ì«‰«  √Ê„— «·‘€·
          ' StrTempFormName = "FrmWorkOrdersData"

        Case ScreensName.ShowPriceScreen
            '‘«‘… ⁄—÷ «·”⁄—
            StrTempFormName = "FrmShowPrice"

        Case ScreensName.TemplateScreen
            '‘«‘… «·⁄—Ê÷ «·Ã«Â“…
            StrTempFormName = "FrmTemplate"

        Case ScreensName.InvoiceScreen
            '‘«‘… ›« Ê—… «·»Ì⁄
            StrTempFormName = "FrmSaleBill"

            '   StrTempFormName = "FrmOut"
        Case ScreensName.PurchaseScreen
            '‘«‘… ›« Ê—… «·„‘ —Ì« 
            StrTempFormName = "FrmBillBuy"

        Case ScreensName.RetrunPurchse
            '‘«‘… ›« Ê—… „— Ã⁄ «·„‘ —Ì« 
            StrTempFormName = "FrmReturnpurchases"

        Case ScreensName.RetrunSalles
            '‘«‘… ›« Ê—… „— Ã⁄ «·„»Ì⁄« 
            StrTempFormName = "FrmReturnSalling"

        Case ScreensName.DestructionScreen
            '‘«‘… «· ·›Ì« 
            StrTempFormName = "FrmDestruction"

        Case ScreensName.OpenStockBalance
            '‘«‘… «·—’Ìœ «·«›  «ÕÏ
            StrTempFormName = "FrmOpeningBalance"
         
        Case ScreensName.StockCountScreen
            '‘«‘… «·Ã—œ
            StrTempFormName = "FrmGard"

        Case ScreensName.StockSettlementScreen
            '‘«‘…  ”ÊÌ… «·„Œ“Ê‰
            StrTempFormName = "FrmStockSettlement"

        Case ScreensName.StockTransfereScreen
            ' ÕÊÌ· „‰ „Œ“‰ ·„Œ“‰
            StrTempFormName = "FrmMoving"

        Case ScreensName.CheckItemQty
            '‘«‘… «·≈” ⁄·«„ ⁄‰ ﬂ„Ì… «·’‰›
            StrTempFormName = "FrmSearchSerial"

        Case ScreensName.CheckItemSerial
            '‘«‘… «·√” ⁄·«„ ⁄‰ ”Ì—Ì«· ·’‰› „⁄Ì‰
            StrTempFormName = "FrmSerialData"

        Case ScreensName.PopUpShowItemCardScreen
            '‘«‘… ⁄—÷  ﬁ«—Ì— ”—Ì⁄… ⁄‰ «·’‰›
            StrTempFormName = "FrmReports"

        Case ScreensName.PopUpShowCustomerBalanceScreen
            '‘«‘… ⁄—÷  ﬁ«—Ì— ”—Ì⁄… ⁄‰ «·⁄„·«¡ «·„Ê—œÌ‰
            StrTempFormName = "FrmSelectDate"

        Case ScreensName.ExpensesTypes
            '‘«‘… √‰Ê«⁄ «·„’—Ê›« 
            StrTempFormName = "FrmExpensesType"

        Case ScreensName.RevenuesTypes
            '‘«‘… √‰Ê«⁄ «·≈Ì—«œ« 
            StrTempFormName = "FrmRevenuesTypes"

        Case ScreensName.ExpensesDataScreen
            '«‰Ê«⁄ «·„’—Ê›« 
            StrTempFormName = "FrmExpenses"

        Case ScreensName.PaymentsDataScreen
            '‘«‘… «·„œ›Ê⁄« 
            StrTempFormName = "FrmPayments"

        Case ScreensName.CashingDataScreen
            '‘«‘… «·„ﬁ»Ê÷« 
            StrTempFormName = "FrmCashing"

        Case ScreensName.AllowsDiscountsScreen
            '‘«‘… «·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…
            StrTempFormName = "FrmDiscounts"

        Case ScreensName.PopUpShowPaymentTime
            '‘«‘… «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…
            StrTempFormName = "FrmPaymentTime"

        Case ScreensName.ReceiptPartScreen
            '‘«‘…  Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ
            StrTempFormName = "FrmReceiptPart"

        Case ScreensName.PopUpShowInstallmentMustPay
            '‘«‘… «·√ﬁ”«ÿ «·„” Õﬁ…
            StrTempFormName = "FrmInstallmentMustPay"

        Case ScreensName.BoxesStockScreen
            '‘«‘… Ã—œ «·Œ“‰…
            StrTempFormName = "FrmBoxStock"

        Case ScreensName.PopUpShowBoxesAccounts
            '—’Ìœ «·Œ“‰… «·√‰
            StrTempFormName = "FrmBoxesAccounts"

        Case ScreensName.BarCodeDesign
            StrTempFormName = "FrmBarcode"

        Case ScreensName.StatisticsShow
            StrTempFormName = "FrmStatistics"

        Case ScreensName.PopUpShowItemCostScreen
            StrTempFormName = "FrmItemCostShow"

        Case ScreensName.OptionsScreen
            StrTempFormName = "FrmOptions"

        Case ScreensName.PopUpShowCusBalances
            '«·≈” ⁄·«„ ⁄‰ √—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰
            StrTempFormName = "FrmShowCusBalances"

        Case ScreensName.PopUpShowItemsRequest
            '«·≈” ⁄·«„ ⁄‰ «·√’‰«› «· Ï »·€  Õœ «·ÿ·»
            StrTempFormName = "FrmRequest"

        Case ScreensName.PopUpShowItemQuantity
            '«·≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›
            StrTempFormName = "FrmSearchSerial"

        Case ScreensName.PopUpShowGuaranteeAlram
            StrTempFormName = "FrmGuaranteeAlram"

        Case ScreensName.PopUpSowStagnantItems
            StrTempFormName = "FrmStagnantItems"

        Case ScreensName.PopUpShowStockMovement
            StrTempFormName = "FrmStockMovement"

        Case ScreensName.ItemsPricePlane
            StrTempFormName = "FrmItemsPrices"

        Case ScreensName.ItemsMainPriceLise
            StrTempFormName = "FrmMainPriceList"
    End Select

    GetFormName = StrTempFormName
End Function

Public Sub OpenScreen(ScreenName As ScreensName, _
                      Optional Lngid As Long = 0, _
                      Optional AnyExtraParm As Variant, _
                      Optional BolPlaySound As Boolean = False, _
                      Optional ExtraParm As Variant, _
                      Optional ExtraParm1 As Variant, _
                      Optional ExtraParm2 As Variant, _
                      Optional OwnerFrm As Form = Nothing)
    
    Dim StrFormName As String
    Dim Msg As String
    Dim Frm As Form
    Dim i As Integer

    On Error GoTo ErrTrap
    StrFormName = GetFormName(ScreenName)

    If StrFormName = "" Then
        MsgBox "OpenScreen:StrFormName"
    End If
 If StrFormName = "FrmSearchSerial1" Then StrFormName = "FrmSearchSerial"
    If StrFormName <> "" Then
        If DoPremis(Do_Open, StrFormName, True) = True Then
            If ScreenName = EmployeesScreen Then
                '»Ì«‰«  «·„ÊŸ›Ì‰
                Load FrmEmployee

                If Lngid <> 0 Then
                    FrmEmployee.Retrive Lngid
                End If

                FrmEmployee.show
                FrmEmployee.ZOrder 0
            ElseIf ScreenName = CustomersScreen Then
                '»Ì«‰«  «·⁄„·« ¡
         
                    Load FrmCustemers
    
                    If Lngid <> 0 Then
                        FrmCustemers.Retrive Lngid
                    End If
    
                    FrmCustemers.show
                    FrmCustemers.ZOrder 0
        
            ElseIf ScreenName = SuppliersScreen Then
                '»Ì«‰«  «·„Ê—œÌ‰
                Load FrmCompany

                If Lngid <> 0 Then
                    FrmCompany.Retrive Lngid
                End If

                FrmCompany.show
                FrmCompany.ZOrder 0
            ElseIf ScreenName = OtherCustomersScreen Then
                '·„ ⁄«„·Ê‰
                Load FrmOtherCustomers

                If Lngid <> 0 Then
                    FrmOtherCustomers.Retrive Lngid
                End If

                FrmOtherCustomers.show
                FrmOtherCustomers.ZOrder 0
            ElseIf ScreenName = ManCompaniesScreen Then
                '»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…
                Load FrmManCompanies

                If Lngid <> 0 Then
                    FrmManCompanies.Retrive Lngid
                End If

                FrmManCompanies.show
                FrmManCompanies.ZOrder 0
            ElseIf ScreenName = ItemsGroupsScreen Then
                '„Ã„Ê⁄«  «·√’‰«›
                Load FrmGroups

                If Lngid <> 0 Then
                    FrmGroups.Retrive Lngid
                End If

                FrmGroups.show
                FrmGroups.ZOrder 0
            ElseIf ScreenName = ItemsDataScreen Then
                '»Ì«‰«  «·√’‰«›
                Load FrmItems

                If Lngid <> 0 Then
                    FrmItems.Retrive Lngid
                End If

                FrmItems.show
                FrmItems.ZOrder 0
            ElseIf ScreenName = StoresDataScreen Then
                '»Ì«‰«  «·„Œ«“‰
                Load FrmStoreData

                If Lngid <> 0 Then
                    FrmStoreData.Retrive Lngid
                End If

                FrmStoreData.show
                FrmStoreData.ZOrder 0
            ElseIf ScreenName = BanksDataScreen Then
                '»Ì«‰«  «·»‰Êﬂ
                Load FrmBanksData

                If Lngid <> 0 Then
                    FrmBanksData.Retrive Lngid
                End If

                FrmBanksData.show
                FrmBanksData.ZOrder 0
            ElseIf ScreenName = BoxesDataScreen Then
                '»Ì«‰«  «·Œ“‰
                Load FrmBoxesData

                If Lngid <> 0 Then
                    FrmBoxesData.Retrive Lngid
                End If

                FrmBoxesData.show
                FrmBoxesData.ZOrder 0
            ElseIf ScreenName = WorkOrdersDataScreen Then
        '        Load frmworkordersdata

        '        If Lngid <> 0 Then
        '            frmworkordersdata.Retrive Lngid
        '        End If
'
'                frmworkordersdata.show
            ElseIf ScreenName = ScreensName.CurrencyDataScreen Then
                '»Ì«‰«  «·⁄„·« 
                Load FrmCurrencyData

                If Lngid <> 0 Then
                    FrmCurrencyData.Retrive Lngid
                End If

                FrmCurrencyData.show
                FrmCurrencyData.ZOrder 0
            ElseIf ScreenName = ScreensName.ShowPriceScreen Then
                '‘«‘… ⁄—Ê÷ «·√”⁄«—
                Load FrmShowPrice

                If Lngid <> 0 Then
                    FrmShowPrice.Retrive Lngid
                End If

                FrmShowPrice.show
                FrmShowPrice.ZOrder 0
            ElseIf ScreenName = TemplateScreen Then
                '‘«‘… «·⁄—Ê÷ «·Ã«Â“…
                Load FrmTemplate

                If Lngid <> 0 Then
                    FrmTemplate.Retrive Lngid
                End If

                FrmTemplate.show
                FrmTemplate.ZOrder 0
            ElseIf ScreenName = InvoiceScreen Then
                '‘«‘… ›« Ê—… «·»Ì⁄
                ' Set Frm = New frmsalebill
                ' Load Frm
                ' If Lngid <> 0 Then
                '     Frm.Retrive Lngid
                ' End If
                ' Frm.Show
                ' Frm.ZOrder 0
           
                ' Set Frm = New frmsalebill
                Load frmsalebill

                If Lngid <> 0 Then
                    frmsalebill.Retrive Lngid
                End If

                frmsalebill.show
                frmsalebill.ZOrder 0
            
            ElseIf ScreenName = PurchaseScreen Then
                '‘«”‘… ›« Ê—… «·„‘ —Ì« 
                '  Set Frm = New FrmBillBuy
                '  Load Frm
                '  If Lngid <> 0 Then
                '      Frm.Retrive Lngid
                '  End If
                '  Frm.Show
                '  Frm.ZOrder 0
     
                Load FrmBillBuy

                If Lngid <> 0 Then
                    FrmBillBuy.Retrive Lngid
                End If

                FrmBillBuy.show
                FrmBillBuy.ZOrder 0
          
            ElseIf ScreenName = RetrunPurchse Then
                '‘«‘… „— Ã⁄ «·„‘ —Ì« 
                Load FrmReturnpurchases

                If Lngid <> 0 Then
                    FrmReturnpurchases.Retrive Lngid
                End If

                FrmReturnpurchases.show
            ElseIf ScreenName = RetrunSalles Then
                '‘«‘… „— Ã⁄ «·„»Ì⁄« 
                Load FrmReturnSalling

                If Lngid <> 0 Then
                    FrmReturnSalling.Retrive Lngid
                End If

                FrmReturnSalling.show
            ElseIf ScreenName = OpenStockBalance Then
                '‘«‘… «·—’Ìœ «·≈›  «ÕÏ ··„Œ«“‰
                Load FrmOpeningBalance

                If Lngid <> 0 Then
                    FrmOpeningBalance.Retrive Lngid
                End If

                FrmOpeningBalance.show
            ElseIf ScreenName = DestructionScreen Then
                '‘«‘… «· ·›Ì« 
                Load FrmDestruction

                If Lngid <> 0 Then
                    FrmDestruction.Retrive Lngid
                End If

                FrmDestruction.show
                FrmDestruction.ZOrder 0
            ElseIf ScreenName = ScreensName.StockCountScreen Then
                '‘«‘… Ã—œ «·„Œ“‰
                Load FrmGard

                If Lngid <> 0 Then
                    FrmGard.Retrive Lngid
                End If

                FrmGard.show
                FrmGard.ZOrder 0
            ElseIf ScreenName = ScreensName.StockTransfereScreen Then
                '‘«‘…  ÕÊÌ· „‰ „Œ“‰ ·„Œ“‰
                Load FrmMoving

                If Lngid <> 0 Then
                    FrmMoving.Retrive Lngid
                End If

                FrmMoving.show
                FrmMoving.ZOrder 0
            ElseIf ScreenName = ScreensName.StockSettlementScreen Then
                '‘«‘…  ”ÊÌ… «·„Œ“Ê‰
                Load FrmStockSettlement

                If Lngid <> 0 Then
                    FrmStockSettlement.Retrive Lngid
                End If

                FrmStockSettlement.show
                FrmStockSettlement.ZOrder 0
            ElseIf ScreenName = ExpensesTypes Then
                '‘«‘… «‰Ê«⁄ «·„’—Ê›« 
                Load FrmExpensesType
                FrmExpensesType.show
                FrmExpensesType.ZOrder 0
            ElseIf ScreenName = RevenuesTypes Then
                Load FrmRevenuesTypes
                FrmRevenuesTypes.show
                FrmRevenuesTypes.ZOrder 0
            ElseIf ScreenName = ExpensesDataScreen Then
                '‘«‘… »Ì«‰«  «·„’—Ê›« 
                Load FrmExpenses2

                If Lngid <> 0 Then
                    FrmExpenses2.Retrive Lngid
                End If

                FrmExpenses2.show
                FrmExpenses2.ZOrder 0
            ElseIf ScreenName = PaymentsDataScreen Then
                '‘«‘… »Ì«‰«  «·„œ›Ê⁄« 
                Load FrmPayments

                If Lngid <> 0 Then
                    FrmPayments.Retrive Lngid
                End If

                FrmPayments.show
                FrmPayments.ZOrder 0
            ElseIf ScreenName = CashingDataScreen Then
                '‘«‘… »Ì«‰«  «·„ﬁ»Ê÷« 
                Load FrmCashing

                If Lngid <> 0 Then
                    FrmCashing.Retrive Lngid
                End If

                FrmCashing.show
                FrmCashing.ZOrder 0
            ElseIf ScreenName = AllowsDiscountsScreen Then
                '‘«‘… «·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…
                Load FrmDiscounts

                If Lngid <> 0 Then
                    FrmDiscounts.Retrive Lngid
                End If

                FrmDiscounts.show
                FrmDiscounts.ZOrder 0
            ElseIf ScreenName = PopUpShowPaymentTime Then

                '‘«‘… «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…
                If ShowCurrencyAlarm(True) = True Then
                    FrmPaymentTime.show
                    FrmPaymentTime.ZOrder 0
                End If

            ElseIf ScreenName = ReceiptPartScreen Then
                '‘«‘…  Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ
                Load FrmReceiptPart

                If Lngid <> 0 Then
                    FrmReceiptPart.Retrive Lngid
                End If

                FrmReceiptPart.show
                FrmReceiptPart.ZOrder 0
            ElseIf ScreenName = PopUpShowCusBalances Then
                Load FrmShowCusBalances
                FrmShowCusBalances.show
            ElseIf ScreenName = PopUpShowInstallmentMustPay Then

                '‘«‘… «·√ﬁ”«ÿ «·„” Õﬁ…
                If ShowInstallmentMustPay(True) = True Then
                    FrmInstallmentMustPay.show
                    FrmInstallmentMustPay.ZOrder 0
                End If

            ElseIf ScreenName = BoxesStockScreen Then
                Load FrmBoxStock

                If Lngid <> 0 Then
                    FrmBoxStock.Retrive Lngid
                End If

                FrmBoxStock.show
            ElseIf ScreenName = PopUpShowItemsRequest Then

                '«·≈” ⁄·«„ ⁄‰ «·√’‰«› «· Ï »·€  Õœ «·ÿ·»
                If ShowRequest(True) = True Then
                    FrmRequest.show
                    FrmRequest.ZOrder 0
                End If

            ElseIf ScreenName = PopUpShowItemQuantity Then
                '«·≈” ⁄·«„ ⁄‰ ﬂ„Ì… ’‰›
                Load FrmSearchSerial
                FrmSearchSerial.show , mdifrmmain
            ElseIf ScreenName = PopUpShowBoxesAccounts Then
                '‘«‘… «·√” ⁄·«„ ⁄‰ «—’œ… «·Œ“‰
                ShowBoxesAccouns
            
            ElseIf ScreenName = CheckItemQty Then
                Load FrmSearchSerial

                If Lngid <> 0 Then
                    FrmSearchSerial.DCboItemsName.BoundText = Lngid
                     FrmSearchSerial.DcboAssbliedItems.BoundText = Lngid
                    FrmSearchSerial.DataCombo1.BoundText = Lngid
                   
                End If

                FrmSearchSerial.Cmd_Click 0
                FrmSearchSerial.show ' vbModal
                
            ElseIf ScreenName = CheckItemswaped Then
                Load FrmSearchSerial1

                If Lngid <> 0 Then
                    FrmSearchSerial1.DcboAssbliedItems.BoundText = Lngid
                End If

                FrmSearchSerial1.Cmd_Click 0
                FrmSearchSerial1.show ' vbModal
                
                
            ElseIf ScreenName = PopUpShowGuaranteeAlram Then
                Load FrmGuaranteeAlram
                FrmGuaranteeAlram.show
            ElseIf ScreenName = CheckItemSerial Then
                Load FrmSerialData
                FrmSerialData.show

                If Lngid <> 0 Then
                    FrmSerialData.DCboItemName.BoundText = Lngid
                End If

                If Not IsMissing(AnyExtraParm) Then
                    FrmSerialData.XPTxtCode.text = CStr(AnyExtraParm)
                    FrmSerialData.Cmd_Click 0
                End If

            ElseIf ScreenName = PopUpShowCustomerBalanceScreen Then
                Load FrmSelectDate

                If Lngid <> 0 Then
                    i = GetDealerType(Lngid)

                    If i = 1 Then
                        FrmSelectDate.CboDealerType.ListIndex = 0
                    ElseIf i = 2 Then
                        FrmSelectDate.CboDealerType.ListIndex = 1
                    ElseIf i = 3 Then
                        FrmSelectDate.CboDealerType.ListIndex = 3
                    End If

                    FrmSelectDate.DcboCusName.BoundText = Lngid
                End If

                If Not IsMissing(AnyExtraParm) Then
                    FrmSelectDate.CboReportType.ListIndex = CInt(AnyExtraParm)
                End If

                If Not OwnerFrm Is Nothing Then
                    FrmSelectDate.show , OwnerFrm
                Else
                    FrmSelectDate.show , mdifrmmain
                End If

            ElseIf ScreenName = PopUpShowItemCardScreen Then
                Load FrmSelectData

                If Lngid <> 0 Then
                    FrmSelectData.DCboItemName.BoundText = Lngid
                End If

                If Not IsMissing(AnyExtraParm) Then
                    FrmSelectData.DcboStores.BoundText = CLng(AnyExtraParm)
                End If

                If Not IsMissing(ExtraParm) Then
                    If Not IsNull(ExtraParm) Then
                        If Not IsEmpty(ExtraParm) Then
                            FrmSelectData.DTPFrom.value = ExtraParm
                        End If
                    End If
                End If

                If Not IsMissing(ExtraParm1) Then
                    If Not IsNull(ExtraParm1) Then
                        If Not IsEmpty(ExtraParm1) Then
                            FrmSelectData.DTPTo.value = ExtraParm1
                        End If
                    End If
                End If

                If Not IsMissing(ExtraParm2) Then
                    If Not IsNull(ExtraParm2) Then
                        If Not IsEmpty(ExtraParm2) Then
                            If FrmSelectData.CboReportType.ListCount > 0 Then
                                FrmSelectData.CboReportType.ListIndex = ExtraParm2
                            End If
                        End If
                    End If
                End If

                If OwnerFrm Is Nothing Then
                    FrmSelectData.show , mdifrmmain
                Else
                    FrmSelectData.show , OwnerFrm
                End If

            ElseIf ScreenName = StatisticsShow Then

                If SystemOptions.SysDataBaseType = AccessDataBase Then
                    Msg = "Â–Â «·√„ﬂ«‰Ì… „ «Õ… ›ﬁÿ ›Ï ‰”Œ… «·‘»ﬂ«  „‰ »—‰«„Ã œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                Load FrmStatistics
                FrmStatistics.show
            ElseIf ScreenName = PopUpShowItemCostScreen Then

                If SystemOptions.SysMainStockCostMethod <> ModernWeightAverage Then
                    Msg = "«·‰”Œ… «·„Œ’’… ·ﬂ ... ·« ” Œœ„ Â–Â «·Œ«’Ì…"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    '     Exit Sub
                End If

                Load FrmItemCostShow

                If Lngid <> 0 Then
                    FrmItemCostShow.DCboItemName.BoundText = Lngid
                    FrmItemCostShow.Cmd_Click 0
                End If

                FrmItemCostShow.show
            ElseIf ScreenName = OptionsScreen Then
                Load FrmOptions
                FrmOptions.show
            ElseIf ScreenName = PopUpSowStagnantItems Then
                If Lngid = 2 Then
                    FrmStagnantItems.Option2.value = True
                    FrmStagnantItems.Option1.value = False
                ElseIf Lngid = 1 Then
                    FrmStagnantItems.Option2.value = False
                    FrmStagnantItems.Option1.value = True
                End If
                Load FrmStagnantItems
                FrmStagnantItems.show
            ElseIf ScreenName = PopUpShowStockMovement Then
                Load FrmStockMovement
                FrmStockMovement.show
            ElseIf ScreenName = ItemsPricePlane Then
               ' Load FrmItemsPrices
               ' FrmItemsPrices.show
            ElseIf ScreenName = ItemsMainPriceLise Then
                Load FrmMainPriceList
                FrmMainPriceList.show
                FrmMainPriceList.ZOrder 0
            End If
        End If
    End If

    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ › Õ «·‘«‘…"
    Msg = Msg & CHR(13) & "Description:" & Err.Description
    Msg = Msg & CHR(13) & "Number:" & Err.Number
    Msg = Msg & CHR(13) & "Source" & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub ShowDialogItemsSearch(m_DataCombo As DataCombo)
    Dim Frm As FrmItemSearch
    Set Frm = New FrmItemSearch
    Frm.RetrunType = 1
    Set Frm.DcboItems = m_DataCombo
    Frm.show vbModal
End Sub
