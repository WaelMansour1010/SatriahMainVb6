Attribute VB_Name = "ModMain"

Option Explicit
'************************
Private Const LOCALE_SDATE                 As Long = &H1D    'date separator
Private Const LOCALE_STIME                 As Long = &H1E    'time separator
Private Const LOCALE_SSHORTDATE            As Long = &H1F    'short date format string
Private Const LOCALE_SLONGDATE             As Long = &H20    'long date format string
Private Const LOCALE_STIMEFORMAT           As Long = &H1003  'time format string
Private Const LOCALE_IDATE                 As Long = &H21    'short date format ordering
Private Const LOCALE_ILDATE                As Long = &H22    'long date format ordering
Private Const LOCALE_ITIME                 As Long = &H23    'time format specifier
 Public APIURL As String
Private Declare Function SetLocaleInfo& Lib "kernel32" Alias "SetLocaleInfoA" (ByVal _
Locale As Long, ByVal LCType As Long, ByVal lpLCData As String)
'************************
Public SystemOptions       As MainOptions
Public allowloadmdifrmmain As Boolean
Public Const SW_SHOWNORMAL = 1
Public HidLowering    As Boolean
Public AllowSelectEmp As Boolean
Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public CurrentVersion            As String

Public NoOFDigitUserTrans        As Integer
Public StoreDigit                As Integer
Public mPosD                     As String
Public mServerD                  As String
    
Public IsSerialByUserTrans       As Boolean
Public ExpensesCoding            As Boolean
Public mZakamsg As String
Public InstallmntsvchrCoding     As Boolean
Public ExpensesCoding2           As Boolean
Public AllowProjectBill2Serial   As Boolean
Public NoOFDigitUserVouc         As Integer
Public Ked_digit                 As Integer
Public JLCodeBasedOnBranch       As Boolean
Public IsSerialByUserVouch       As Boolean

Public POSConnection             As New ADODB.Connection
Public ServerDb                  As String
Public POSDb                     As String

Public BranchDigit               As Integer
  
Public SysSQLServerType          As Integer
Public SysSQLServerName          As String
Public SysSQLServerTypeTechnical As String
'Public StrAppRegPath As String
Public SysSQLServerDataBaseName  As String
Public SysSQLServerUserId        As String
Public SysSQLServerUserpassword  As String
 
Public MainBranch                As String
Public MainBranchID              As Long
Public CountAllBranch            As Long
Public CountAllServer            As Long
'-------------------------------
Public MainServer                As String
Public CurrentServer             As String
Public MainServerID              As Long
Public CurrentServerID           As Long
    
'this function to get handle for DeskTop
Public Declare Function GetDesktopWindow _
   Lib "user32" () As Long

'to find any window
Public Declare Function FindWindow _
               Lib "user32" _
               Alias "FindWindowA" (ByVal lpClassName As String, _
                                    ByVal lpWindowName As String) As Long

Public Declare Function GetWindowRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Public Declare Function GetWindowLong _
               Lib "user32" _
               Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

Public Declare Function GetClientRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Public Declare Function InvalidateRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT, _
                             ByVal bErase As Long) As Long

Public Declare Function ClientToScreen _
               Lib "user32" (ByVal hWnd As Long, _
                             lpPoint As POINTAPI) As Long

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long
 
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_LAYOUTRTL = &H400000

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function GetCursorPos _
   Lib "user32" (lpPoint As POINTAPI) As Long

Public Const HWND_TOPMOST = -1

Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOSIZE = &H1

Public Const SWP_NOMOVE = &H2

Public Const SWP_NOACTIVATE = &H10

Public Const SWP_SHOWWINDOW = &H40

Public Declare Sub SetWindowPos _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long)

'This API to show a form
Public Declare Function ShowWindow _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal nCmdShow As Long) As Long

Public Const SW_SHOW   As Long = 5

Public Const SW_SHOWNA As Long = 8
'Public Cn As New ADODB.Connection

Private Const STATE_SYSTEM_FOCUSABLE = &H100000

Private Const STATE_SYSTEM_INVISIBLE = &H8000

Private Const STATE_SYSTEM_OFFSCREEN = &H10000

Private Const STATE_SYSTEM_UNAVAILABLE = &H1

Private Const STATE_SYSTEM_PRESSED = &H8

Private Const CCHILDREN_TITLEBAR = 5

Private Type TITLEBARINFO
    cbSize As Long
    rcTitleBar As RECT
    rgstate(CCHILDREN_TITLEBAR) As Long
End Type

Private Declare Function GetTitleBarInfo _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByRef pti As TITLEBARINFO) As Long

'this api used in the MouseDown_Event to enable the user From
'Move the form From any Postion ...
'Try this with MDI Form by Move it From any Position by the Mouse...
Public Declare Function ReleaseCapture _
   Lib "user32" () As Long

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

Public Const WM_NCLBUTTONDOWN = &HA1    'used in the  Form Mouse Down event

Public Const HTCAPTION = 2              'used in  Form Mouse Down event

Public Const SysMaronColor = &H80&

Public Const SysBackColor = &HE2E9E9

Public Enum SystemInterface
    ArabicInterface
    EnglishInterface
End Enum

Public Enum FormSizeType
    NoChangeInSize
    TransactionSize
    ReportSize
End Enum

Public Enum Versions
    DemoVersion
    RegisterVersion
End Enum

Public Enum WindowState
    NormalWindow
    MinimizedWindow
    MaximizedWindow
    CustomeWindow
End Enum

Public Enum SystemTargets
    ToMahfooz
    ToGamal
    ToNour
    ToBahrin
    ToRashid_Saif
    ToTalal
    ToSobhi
    ToMoains
    ToIbrahimShakr
End Enum

Public Enum StockCostType
    LastPurPriceType = 0 '√Œ— ”⁄— ‘—«¡
    LastSalesPriceType = 1 '√Œ— ”⁄— »Ì⁄
    FirstInFirstOut = 2 '«·Ê«œ— «Ê·« Ì’—› «Ê·«
    LastInFirstOut = 3 '«·Ê«—œ «ŒÌ—« Ì’—› √Ê·«
    WeightAverage = 4 '«·„ Ê”ÿ «·„—ÃÕ ⁄·Ï «·› —…
    ModernWeightAverage = 5 '«·„ Ê”ÿ «·„—ÃÕ »⁄œ ﬂ· ›« Ê—…
End Enum

Private Type MsgData
    MsgID As Integer
    MsgArabic As String * 255
    MsgEnglish As String * 255
End Type

Public Enum ConnectTypes
    ConnecLocal
    ConnectRemote
End Enum

Public Enum ReportDisplayDataType
    DetailDisplayType
    ShortDisplayType
End Enum

Public Enum AppDataBaseTypes
    AccessDataBase
    SQLServerDataBase
End Enum

Public Enum SQLServersTypes
    NotSet '·„ Ì „  ÕœÌœ ‰Ê⁄ «·”Ì—›—
    LocalServer '«·”Ì—›— „Õ·Ï ( ⁄·Ï ‰›” «·ÃÂ«“ «·„ÊÃÊœ ⁄·ÌÂ «·»—‰«„Ã)
    RemoteServer '«·”Ì—›— „ÊÃÊœ ⁄·Ï ÃÂ«“ ›Ï «·‘»ﬂ…
End Enum

Public Enum SQLServersTypesTechnical
    
    server1  '«·”Ì—›— „ÊÃÊœ ⁄·Ï ÃÂ«“ ›Ï «·‘»ﬂ…
    Server2
End Enum

Public Enum TakeDateTypes
    InvDateFromLocalCompuer
    InvDateFromLastInvDate
    InvDateFromServerComputer
End Enum

Public Type RegsAppSections

    FormsSetting As String
End Type

Public Enum AppAccoutingType
    SimpleAccoutning
    CompeleteAccounting
End Enum

Public Type RegisterySections

    DockingPanesSection As String
End Type

Public Enum UserTypes
    UserNourCo '«·„»—„Ã «Ê «·‘—ﬂ… √⁄·Ï ’·«ÕÌ…
    UserAdminAll '„œÌ— «·‰Ÿ«„ «·⁄«„ √⁄·Ï ’·«ÕÌ… ›Ï «·»—‰«„Ã
    UserAdmin '„” Œœ„ „œÌ—
    UserNormal '„” Œœ„ ⁄«œÏ
End Enum

Public Type MainOptions
    SysDataBaseType As AppDataBaseTypes
    SysDataBasePath As String
    SysSQLServerType As SQLServersTypes
    SysSQLServerTypeTechnical As SQLServersTypesTechnical
    
    SysSQLServerUserId As String
    SysSQLServerUserpassword As String
    SysSQLServerName As String
    SysSQLServerDataBaseName As String '«”„ ﬁ«⁄œ… «·»Ì«‰«  œ«Œ· «·”Ì—›—
    SysRegsAppPath As String '„”«— Õ›Ÿ ≈⁄œ«œ«  «·»—‰«„Ã ›Ï «·—ÌÃ” —Ï
    SysPublicUnloadStatus As Boolean
    SysVersion As Versions
    SysServerIP As String
    SysRegFilePath As String
    
    SysRegTempFilePath As String
    SysRegisterState As ExireTypes
    SysRunNumber As Integer
    SysHelp As New HTMLHelp
    SysCurrentAccountIntervalID As Long  'ﬂÊœ «·› —… «·„Õ«”»Ì… «·Õ«·Ì…
    SysInvDateTakeType As TakeDateTypes
    SysPurDateTakeType As TakeDateTypes
    SysCashDateTakeType As TakeDateTypes
    
    SysMantainceAllow As Boolean
    SysAllowStockNegative As Boolean '«·”Õ» ⁄·Ï «·„ﬂ‘Ê› „‰ «·„Œ“‰
    NotAllowStockNegativeInternal As Boolean
    MustEnterNewNo As Boolean
    SysAllowBoxNegative As Boolean '«·”Õ» ⁄·Ï «·„ﬂ‘Ê› ›Ï «·Œ“‰…
    SysAppAccoutingType As AppAccoutingType
    SysMainStockCostMethod As StockCostType 'ÿ—Ìﬁ… Õ”«» «· ﬂ·›… ··„Œ“Ê‰
    SysDefCurrencyForamt As String '«· ‰”Ìﬁ «·≈› —«÷Ï ··⁄„·…
    SysDefQuantityFormat As String '«· ‰”Ìﬁ «·≈› —«÷Ì ··ﬂ„Ì« 
    SysDefQuantityDecimal As Integer '«· ‰”Ìﬁ «·≈› —«÷Ì ··ﬂ„Ì« 

    usertype As UserTypes
    UserShowToolTip As Boolean '⁄—÷  ·„ÌÕ »«·‰”»… ··„” Œœ„
    UserInterface As SystemInterface 'Ê«ÃÂ… «· ÿ»Ìﬁ «·Œ«’… »«·„” Œœ„
    UserWindowState As WindowState '" ( Õ«·… «·‘«‘…) 'ŒÌ«— Œ«’ »«·„” Œœ„ «·Õ«·Ï
    UserInvoiceChangePrice As Integer
    UserInvoiceChangePrice1 As Integer
    UserInvoiceChangePrice2 As Integer
    FixedCustomer  As Integer
    ShowBillCommisions As Integer
     
    AllowItemByRowMove   As Boolean
    AllowItemByRowOut   As Boolean
     
    '31032017egypt
    AllowChangeUnitIqar As Boolean
    AllowCreditPass As Boolean
    AllowProductOrderOne As Boolean
    AllowSalesSaveWithoutCostPrice As Boolean
    AllowChanProjectBillPrice As Boolean
    AllowSalesMultyPayed As Boolean
    AllowPurchasesMultyPayed As Boolean
    AllowDynamicEdit  As Boolean
    AllowScInterface As Boolean
    AllowScInterface2  As Boolean
    ShowOnlyItemsOfSales As Boolean
    GeneralVoucherCreateSalesGE As Boolean
    SalesNotCreateGe As Boolean
    PrintInvoiceByBranch As Boolean
    LinkSupplerWithItem As Boolean
    IsInternalMultiOrder As Boolean
    IsBlue As Boolean
    IsBluee As Boolean
    ApplyEinvoice As Boolean
    CanUploadZakatOpt As Boolean
    IsCahngeServiceInvoice As Boolean
    ServerNameW As String
    DbNameW As String
    ApplyEinvoiceWithActive As Boolean
    ApplyEinvoiceWithBranch  As Boolean
    HiddenBalanceFromBox As Boolean
    EmpAccountByDep As Boolean
         
    Isthickness As Boolean
    IsMashghal  As Boolean
             
    IsSalesOrder As Boolean
    IsQrCodePrint As Boolean
    IsShowItemsBranch As Boolean
    IsElecWaterCont As Boolean
    IsDogeMode As Boolean
    IsMaintItemMode As Boolean
    IsHiddenTransportInv As Boolean
    IsHeaderPrint As Boolean
    HideInfroCasher As Boolean
    CaNUpdateApprovedDoc As Boolean
    CaNUpdateAutoSalesInvoice As Boolean
    OpenAccountAqar As Boolean
    IsMultiItemsInCompItem As Boolean
    LimitDefaultCredit As Double
    LimitDefaultCreditDays As Double
    AllowEditCreditLimit As Boolean
    AllowEditCreditBalance As Boolean
    ProvisionsByıEQuipments As Boolean
    ReturnSAlesByBarcode As Boolean
    CreatePayOrderSales As Boolean
    TripnotUploadExpenses As Boolean
    ExpensesByQtyOnly As Boolean
    DiscountByQtyOnly As Boolean
    IsTransferByCode As Boolean
    ZacatHandW As Boolean
    ShowPrinterDialoge As Boolean
    ShowPrinterDialoge2 As Boolean

    DontDistributeLegalACC As Boolean
    IsBarCodeByUnit As Boolean

    ProvisionsByManagement As Boolean
    AllowEditInvoiceNoticeDiscount As Boolean
    AllowEditInvoiceOfReturn As Boolean

    AllowConvertAlertToJob As Boolean
    ShowBalanceOfEmpInSalary As Boolean
    PaymentIntoAccouStat As Boolean
    CreateJLEmpCommissions As Boolean
    TypeContractAutoFromIqar As Boolean
    AllowRepeatInvoiceNo As Boolean
    EmpSalaryDigts As Integer
    AllowReturnFIFO As Boolean
    AllowDiscountAllowedFIFO As Boolean
    AllowJLManualFIFO As Boolean

    
    
    
    
    


 OpenVATAccountOwner As Boolean
    LinkCustomerWithCars As Boolean
AllowEditCashingLinkProj As Boolean
TransBillPriceByGrid As Boolean
NoCreatJLInRentContract As Boolean
Commonname   As String
SerialNumber   As String
OrganizationName   As String
 

Invoicetype   As Integer
DefaultInvoicetype As Integer
SendingMode   As Integer

 
industrey   As String
CSR   As String
Privatekey   As String
PublickeycertPem   As String
SecretKey   As String

    DealingWithPrepayAccount As Boolean
    NotAllowedCalcVata As Boolean
    AllowSkipDiscountGroup As Boolean
    OpenAtProduction As Boolean
    NotEditInternalPrice As Boolean
    NotEditSalesRetPrice As Boolean

    IssueVoucherWorkWithRemain As Boolean
    TripDateInsertDefulat As Boolean
    TripwithorderOnly As Boolean
    AllowPriceWithWidth As Boolean

    CreateJLVactionAratha As Boolean
    PriceWithVAT As Boolean
    CustomerRecordNoIsnotManda As Boolean
    
    ProjectInvoiceAnalysisJL As Boolean
    AllowWorkCustomerPoints As Boolean
    
    AllowAnalyticJL As Boolean
    InsuranceOnOwner As Boolean
    ServicesOnOwner As Boolean
    AllowCraeJLQuality As Boolean
    CantWorkwithComponenetinEmpScr As Boolean

    AllowChangePriceApprove As Boolean
    AllowSkipPayment As Boolean

    DueComm As Boolean
    DueWater As Boolean
    DueElectr As Boolean
    DueService As Boolean
    CommissionOnOwner As Boolean

    CommissionDue As Boolean
    SupplierReciveGE As Boolean

    SalaryJLByManagement As Boolean
    SendToAprovedSalesBill As Boolean
    AllowAprovedSalesBill As Boolean
    SalaryJLByAnalyEqup As Boolean
    BranchmustimSalary As Boolean

    AllowGoodPerfAccount As Boolean
    ManualSalesInvoiceMust As Boolean
    AllowItemByRow   As Boolean
    AllowChangManualQtyMix   As Boolean
    AccountAccordingCash    As Boolean

    ProductionRawMaterMix   As Boolean
    AllowLastPrice   As Boolean
    AllowNoRoudProjectInvoices  As Boolean
    CountPrint As String
    NOOFPRINTCOPIESSALES As Integer
    AllowDynamicAutoSus As Boolean
    AllowUnbalncedByBranchAccount As Boolean
    AllItemInVAT As Boolean
    CloseMovingVchrinSales  As Boolean
    CantChangeSalesPerson As Boolean
    BatchCreateManyworkOrder As Boolean
      
    AllowProjectBill2Serial As Boolean
    CashCustomerNameMustenter As Boolean
    AllowChangeSalesAtTransfer As Boolean
    SalesTrustsAffectVedorCode  As Boolean
    CanChangeStatusDateRequest As Boolean
    CanChangeTripAfterInvoiceing As Boolean
    DontShowMoreDetailsCompItem As Boolean
    traveDiscountFromCustomerDirect As Boolean
    TransferNotInvItemDef As Boolean
    CanTransferItemDef As Boolean
 
    CustMobNoMandatory As Boolean
    CustVatNoMandatory As Boolean

    '31032017egypt
    'modmod
    CanCustomerandVendor As Boolean
    NotEditDiscountLine As Boolean
    CanOpenWorkOrder As Boolean
    CanChangePriceUpOnly As Boolean
CanProjectAccountOnly As Boolean
CanUploadZakat As Boolean
IsHiddenUser As Boolean
CanPostPumpInv As Boolean
    CanEditMinRentValue As Boolean
    CanAcreditRsContract As Boolean
    CanIsShamel As Boolean
    CanEditLegalAffairs As Boolean
    'OPenShortInvoice As Boolean
                
    OPenShortInvoice As Boolean
      OPenShortInvoicePump As Boolean
        OPenShortInvoicePetrol As Boolean
    
   
  
    MonyeIssueVchrNoMust As Boolean
    POMustentryAndBillMustEntry As Boolean
    USERautoIssueVoucher As Boolean
    HideTbarInPos As Boolean

    IsShowLensesDetails As Boolean
    CanEditOnlyPayMethod As Boolean

    SortInvoiceByEntry As Boolean
    CostProductOrderByOut As Boolean
    SAveInhomePath As Boolean
    CanPartialpayment As Boolean
    EndRentifPayed As Boolean
    cantCahngeAkarinExpenses  As Boolean
    EmployeeSalaryBYBranch  As Boolean
    returnnotcreatvoucher As Boolean
    OnlyOneCashingVchr As Boolean
    CheckDateFormatCorrect As Boolean
    CheckMobileFormatCorrect As Boolean
             
    CantRepetttransferNoforCashing As Boolean
    WaiverSetByContract As Boolean
    NoBooking As Integer
    MultyStore As Boolean
    RawMaterMix2 As Boolean 'modmod
    DontCreateOut As Boolean 'modmod
    DontCreateOut2 As Boolean 'modmod
    InsertItemManualOut As Boolean 'modmod
     
    UserInvoiceShowProfit As Integer
    UserItemsPremis As UsersPremis
    UserScreenPremis As UsersPremis
    UserStorePremis As UsersPremis
    UserAccountsPremis As UsersPremis
    HaveTaxOnSalles As Boolean 'Õ”«» ÷—«∆» „»Ì⁄«  ⁄·Ï «·›Ê« Ì—
    SysConnectionType As ConnectTypes
    SysTarget As SystemTargets
    BolUpdateTaskInProgress As Boolean
    BolStopUpdateTask As Boolean
    Items_or_operation As Integer
    ProjectDiscountPolicy As Integer
    gldetails_or_gl_general As Integer
    ProcessPeriodType As Integer
    
    EmpComponentDigts As Integer
    ImagesPath As String
    Reportpath As String
    BigUserPw As String
     BigUserPw2 As String
    
    itemcodePart1 As Integer
    ChasingStatus As Integer
    itemcodePart2 As Integer
    itemcodePart3 As Integer
    itemcodeSeperator1 As String
    itemcodeSeperator2 As String
    ViewAccountsbyBranch As Boolean
    AllowEditeAccounts As Boolean
    AllowHideAssest As Boolean
    LockSalary As Boolean
    AllowAccountMultyPayed As Boolean 'true
    itemcodePart1NoOFDigit As Integer
    itemcodePart2NoOFDigit As Integer
    itemcodePart3NoOFDigit As Integer
    itemsWorkWithSize As Boolean
    workWithBarcode As Boolean
    WorkWithBarCodeParent As Boolean
    WorkWithLINKEDiTEMS As Boolean
    WorkWithLINKEDiActivity As Boolean
    amlaketbatrentOnly As Boolean
    WorkWithBranchLogo As Boolean
    WorkWithFirstInstallOnly As Boolean
    CreateInsuranceAccountForCustomers As Boolean
    WorkWithGroupCode As Boolean
    DecideItemName As Boolean
    DefaultIsCreditSales As Boolean
    DefaultIsCreditPurchase As Boolean
    DefaultIsCreditPurchaseRet As Boolean
    
    EmpNotExcceedDiscount As Boolean
    BoxLossandIncreae As Boolean
    attacheditemsisfree As Boolean
    EnableCustomerAging As Boolean
    returnByBarCodeOnly As Boolean
    showcostColorininvoice As Boolean
    SubContactorHave3Account As Boolean
    ProjectUnderImplemen As Boolean
    ProjectEmployeeGV As Boolean
    PursgaseWithoutDecimal As Boolean
    workWithCustomerContract As Boolean
    workWithvendorContract As Boolean
    PoCreateVoucher As Boolean
    DiscountSalesCreateVchr As Boolean
    AllowCostPerStore As Boolean
    AllowCostnNewShape As Boolean
    AllowCostBySerial As Boolean
    PaymentDifferent As Boolean
    cancellAllApprove As Boolean
    poWithatotalQty As Boolean
    AnalyticPaymentVouchr As Boolean
    ShowDriverOnly As Boolean

    PayrollOneAccount As Boolean
    WorkWithItemsDetails As Boolean
    FAAddtionCreateAccount As Boolean
    Create2account4Supp As Boolean
    workwithticketAllocation As Boolean
    JLCodeBasedOnBranch As Boolean
    TradingPOS As Boolean
    posshape2 As Boolean
     
    SellOrderBalance As Boolean
    CanChanegeLinkedPurcahsenvoice As Boolean
    CanChanegeLinkedSsalesnvoice As Boolean
    updatecashvchrifdeposite As Boolean
    Revenueowed As Boolean
    AllowupdateJobStatus As Boolean
    OpeningEmployeeShowAll As Boolean
    AllowTowShift As Boolean
    AllowItemsShortName As Boolean
    EndServiceMore5Year As Boolean
    VacstionShowOldSalaries As Boolean
    AllowReturnWithoutCost As Boolean
    ShowItemByCustomer As Boolean

    Ecnomy As Boolean
    WebAdv As String
    DuplicateitemsNames As Boolean
    CostStarting As Boolean
   CostStartingGard As Boolean
    
    chkuserCode As Boolean
    Itemsattachedzero As Boolean
    itemsWorkWithColor As Boolean
    itemsWorkWithDates As Boolean
    itemsWorkWithClass As Boolean
    hidecolumn As Boolean
    ExceedShipment As Boolean
    AllowSett As Boolean
    AllowSett1 As Boolean
    Allowpayroll As Boolean
    AllowCreateHajomraVoucher As Boolean
    AllowRequestgl As Boolean
    Allowrank As Boolean
    AllowCompChanPrice As Boolean
    AllowBigAccount As Boolean
    AllowOrbonDate As Boolean
    AllowShowAllEmployee As Boolean
    DateCanNotEdit As Boolean
    BranchCanNotEdit As Boolean
    PreFixCanNotEdit As Boolean
    AllowPOSPAy As Boolean
    RawMaterMix As Boolean 'modmod
    VATNoAccordActivity  As Boolean
    NotCrtResvVouchProjects  As Boolean
    EmpProduction As Boolean
    ItemProduction As Boolean
    ExpProduction As Boolean
    InvoiceTransferJLTotal As Boolean
    CarsRevenuePerOwner As Boolean
    IsCustSalesManCashRelated As Boolean
    showEmployeeAccountIntrip As Boolean
    DUEDOCUMENTbyinstallDate  As Boolean
    CanSkipPurchOrder As Boolean
    CompilingBasedTable As Boolean
    DontSaveInvoiceWithoutDocType As Boolean
    DontDuplicateManulaNoInPurchase As Boolean
    SpecialVersion As Boolean

    CanEditCars As Boolean

    CanChangeOut As Boolean
    CanCancelContract As Boolean

    AllowSaveTripWithoutExpen As Boolean
    SAVEMAINTENANCEJOBWITHORDERORPLANONLY As Boolean
 CustCreat4Acc As Boolean
 SuppCreat4Acc As Boolean
 CreateEntryBillItems As Boolean
    '  AllowAccountMultyPayed As Boolean
    HideCost As Boolean
    LinkUsersWithPayment  As Boolean
    ItemcodeGroupOnly As Boolean
    SaleDiscount1 As Integer
    SaleDiscount2 As Integer
    SaleDiscount3 As Integer
    SaleDiscount4 As Integer
    autoIssueVoucher  As Boolean
    MonthIs30days As Boolean
    autoReseiveVoucher As Boolean
    ReturnSallingOption As Boolean
    Ked_digit As Integer
    Count_ACCOUNT_digit As Integer
    Save_options As Integer
    bankComm As Boolean
    ChequeBox As Boolean
    IsCheque As Boolean
    CustomerhavethreeAccounts As Boolean
    IsCreateOpenBalnceMan As Boolean
    CustomerhavethreeAccounts1 As Boolean
    CostByProduction As Boolean
    IsByNewCoding As Boolean

    IsAutoNameItems As Boolean
    mDomainData As String
    AllowRepeateCar As Boolean
    CanPrintMultiSales As Boolean
    CanPayWithoutPrint As Boolean
    PlaywithAuthorityMatrix As Boolean
    AllowEditProductionOutManulay As Boolean
    AllowEditVaTManulay As Boolean
    ShowOldAccountReports As Boolean
    MaintOrderCantRepeatSales As Boolean
    MaintOrderCantRepeatBillBuy As Boolean
    TripRevenueAuto As Boolean
    cdoSMTPServer  As String
    TxtFromName  As String
    txtFromEmail  As String

    cdoSendUserName  As String
    cdoSendPassword  As String
    cdoSMTPUseSSL As Boolean

    cdoSMTPServerPort As Integer

    PaymentMethLaterCompItem As Boolean
    ShowBalanceCustInv As Boolean
    IsSerialByUserTrans As Boolean
    IsSerialByUserVouch As Boolean
    NoOFDigitUserTrans As Integer
    NoOFDigitUserVouc As Integer
    IsSomeItemWeight As Boolean
    IsMergeVat As Boolean
    
    FromNo  As Double
    OrNo  As Double
    CodeFrom  As Double
    CodeTo  As Double
    WeightFrom  As Double
    WeightTo  As Double
             
    IsGeometricProportions As Boolean

    logowidth As Double
    logoHeight As Double
    CreateDriverBox As Boolean
    CreateDriverEra As Boolean

    TypicalProduction As Boolean
    ExpensesCoding As Boolean
    ExpensesCoding2 As Boolean
    SMSUserName As String
    SMSPassWord As String
    SenderName As String
    OPTWEB As Integer
    CLockedDate As Date
    Alarm_start1 As Date
    LockSystem As Double
    InstallmntsvchrCoding As Boolean
    AllowIndirectCost As Boolean
    AssetAccount As Boolean
    AssetAccount1 As Boolean
    banks_Accounts3 As Boolean
    ReturnSallingIntervalCount As Integer
    DateOpt As Integer
    ReturnSallingIntervalCount1 As Integer
    IndirectCostPercentage As Double
    StoreDigit As Double
    BranchDigit As Double
    AllowCommtionJEFromValueVisa As Boolean
    AllowWorkWithArea As Boolean
    AllowAcceleratepayment As Boolean
    ItemcodeGroupandParentGroup As Boolean
    ReservEmp As Integer
    itemSeprator As String '    ›«’·  ﬂÊÌœ «·«’‰«›
    DefaultQtyTrans As Double
    StoreAccountHaveSettelment As Boolean
    eachStoreHaveLossAccount As Boolean
    eachStoreHaveGiftAccount As Boolean
    AllowExperDateFIFO As Boolean
End Type

'Public FrmNewsBarPane As FrmPane '‘—Ìÿ «·√Œ»«— Ê«·„⁄·Ê„« 

'Public FrmOutBarPane As FrmOurBarPane '‘—Ìÿ «·√Œ ’«—« 

'Public ItemsTreePane As FrmPaneTree '‘Ã—… «·√’‰«›

'Public FrmMantaincePane As FrmPane '‘—Ìÿ «·’Ì«‰…

'Public FrmInternetNews As FrmPane '√Œ»«— «·√‰ —‰ 

Public FrmDynamicHelpPane As FrmPaneHelp '‰«›–… «·„”«⁄œ… «·›Ê—Ì…

'Public FrmCalendarPane As FrmPaneCalendar

Public Enum DockingPanesIDs
    OutBarPaneID = 1
    NewsBarPaneID = 2
    ItemsTreeID = 3
    MantainceID = 4
    InternetNews = 5
    DynamicHelp = 6
    CalendarPaneID = 7
End Enum

Public OPEN_NEW_SCREEN As Boolean

Public Decimal_Places  As Integer

Public Decimal_Places1 As Integer

Public PrintBranchINGE As Boolean

Public PrintCCinGE     As Boolean

Public ChartPrintinAS  As Boolean

Public HideAllAlarms   As Boolean

Public Messnger        As Boolean
Public AlarmAuto       As Boolean
Public ViewAging       As Boolean
Dim sql                As String
Public Declare Function InternetGetConnectedState _
               Lib "wininet" (ByRef dwFlags As Long, _
                              ByVal dwReserved As Long) As Long
Public Const MsssageSeconde = 2
Public messageResult As String

Public Sub OpenWebSite(Optional StrURL As String = "")

    If StrURL = "" Then
        StrURL = "www.sattary.com"
    End If
    Dim r As Long
    ' r = ShellExecute(0, "open", StrURL, 0, 0, 1)
   
    If InternetGetConnectedState(0, 0) = 1 Then
        Shell "explorer " & StrURL & "", vbMinimizedFocus
        'MsgBox "Connected"
   
    End If

End Sub

Private Function CreateKey() As String
    Dim Msg As String
    Msg = GetHardDiskData(False)
    Msg = Msg & "**" '& GetProcessorData(False)

    Dim i         As Integer
    Dim StrOutKey As String
    Dim StrChar   As String
    Dim LngOutKey As Long
    Dim VarTemp   As Variant
    VarTemp = Split(Msg, "**", , vbBinaryCompare)

    For i = 1 To Len(Msg)
        StrChar = mId(Msg, i, 1)

        If StrChar <> "" Then
            LngOutKey = LngOutKey + (Asc(StrChar) * 12)
        End If

    Next i

    LngOutKey = (LngOutKey + LngOutKey)
    LngOutKey = LngOutKey * 3
    CreateKey = CStr(LngOutKey)
End Function

Public Sub Main()
    allowloadmdifrmmain = False
    Dim AskOption      As Boolean
    Dim Msg            As String
    Dim BolShowRequest As Boolean
 
    On Error GoTo ErrTrap
    '-----------------Set Main Options------------------
    'Change System Format
      Dim lcid As Long
    'Call SetLocaleInfo(lcid, LOCALE_SSHORTDATE, CStr("dd/MM/yyyy"))
    Call SetLocaleInfo(lcid, LOCALE_SSHORTDATE, CStr("yyyy/MM/dd"))
 
    '*****************************************
    
    SystemOptions.SysAppAccoutingType = CompeleteAccounting
    SystemOptions.SysDataBaseType = SQLServerDataBase
    SystemOptions.SysSQLServerDataBaseName = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")
    SystemOptions.SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
    SystemOptions.SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")
    
    SystemOptions.SysTarget = ToNour
    SystemOptions.UserInterface = ArabicInterface
    SystemOptions.SysConnectionType = ConnectRemote
    SystemOptions.SysMantainceAllow = True
 
    StrAppRegPath = "bisegypt\SimpleAccounting"
    SysSQLServerType = val(GetSetting(StrAppRegPath, "ServerCon", "ServerType", 0)) '0 loca 1 not 2 rem
    SysSQLServerName = GetSetting(StrAppRegPath, "ServerCon", "ServerName", "")
    SysSQLServerTypeTechnical = GetSetting(StrAppRegPath, "ServerCon", "SysSQLServerTypeTechnical", "0")

    SysSQLServerDataBaseName = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")

    SysSQLServerUserId = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserId", "salim")
    SysSQLServerUserpassword = GetSetting("Byte_DBS", "Setting", "SysSQLServerUserpassword", "salim")
 
    '---------------Set The Application Title-----------

    App.Title = GetAppTitle
    '---------------------------------------------------
    LoadSettings
    'Dim x As Integer
    'Dim dateDifferent As Integer
    'If Alarm_start <> "" Then
    '
    'dateDifferent = DateDiff("d", Date, Alarm_start)
    '
    'If dateDifferent <= 0 Then
    '
    '        x = DateDiff("d", Date, Alarm_end)
    '        If x <= 0 Then
    '             MsgBox "ÌÃ» „—«Ã⁄Â «·‘—ﬂ… ·«⁄«œ…  ‘€Ì· «·»—‰«„Ã", vbInformation: End
    '        Else
    '            MsgBox "ÌÊÃœ ⁄·Ìﬂ ﬁ”ÿ „” Õﬁ ⁄‰ ﬁÌ„… «·»—‰«„Ã »—Ã«¡ ”œ«œ… ﬁ»·  " & x & "   ÌÊ„", vbCritical
        
    '        End If
    'End If
    '
    'End If

    'If SystemOptions.CLockedDate < Date Then
    '    Msg = "«·»—‰«„Ã ·« Ì” ÿÌ⁄ «·« ’«· »ﬁ«⁄œ… «·»Ì«‰«  —Ã«¡ «·« ’«· »«·œ⁄„ «·›‰Ï...!!!"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    'End
    '   Cn.Execute "update notes_all  set LockedDate='" & SQLDate(Date) & "'"
    
    'End If

    '---------------Set The Main Application Vatiables--
    StrAppRegPath = "bisegypt\SimpleAccounting"
    SystemOptions.SysRegsAppPath = StrAppRegPath
    SystemOptions.SysHelp.CHMFile = App.path & "\HelpFiles\HelpFiles.chm"
    App.HelpFile = App.path & "\HelpFiles\HelpFiles.chm"
    SystemOptions.SysServerIP = "BYTE"
    SystemOptions.SysRegFilePath = App.path & "\RegFile.txt"
    SystemOptions.SysRegTempFilePath = App.path & "\TempRegFile.txt"
    '---------------------------------------------------
    CreatLogFile
   
    GoTo ll

    If State = "" Then 'first run
        save_confoguration "X0569220500", 1, CreateKey + "10111982"
        'FrmRegisteration.LblExpireCount = 50
        'FrmRegisteration.Show vbModal
        SystemOptions.SysRegisterState = DemoRun

        'Exit Sub
    Else

        If State = "X0569220500" And run_count < 50 Then 'demorun
            save_confoguration "X0569220500", run_count + 1, CreateKey + "10111982"
            'MsgBox "hi"
            'SystemOptions.SysRegisterState = CheckExpireation
            'FrmRegisteration.Show
            'FrmRegisteration.LblExpireCount = run_count
            SystemOptions.SysRegisterState = DemoRun
            'Exit Sub
        Else

            If State = "X0569220500" And run_count = 50 Then 'dem stop
                MsgBox " „ «‰ Â«¡ «·‰”Œ… «· Ã—Ì»Ì… « ’· »«·‘—ﬂ… ›Ì Õ«·… «·—€»… ›Ì «·«” „—«— Ê«·Õ’Ê· ⁄·Ï „› «Õ «·Õ„«Ì…", vbCritical
                'FrmRegisteration.Show
                'FrmRegisteration.LblExpireCount = run_count
                SystemOptions.SysRegisterState = DemoStop
                'Exit Sub
            Else

                If State = "D11002D19y84" And key_for_me = CreateKey + "10111982" Then 'registered
                    SystemOptions.SysRegisterState = Registered
                Else

                    If State = "D11002D19y84" And key_for_me <> CreateKey + "10111982" Then ' play in registry hack
                        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì „·› «·Õ„«Ì… « ’· »„œÌ— «·‰Ÿ«„ ·Õ· «·„‘ﬂ·… Ê «·Õ’Ê· ⁄·Ï «·„› «Õ «·’ÕÌÕ", vbCritical
                        'FrmRegisteration.Show
                        'FrmRegisteration.LblExpireCount = run_count
                        SystemOptions.SysRegisterState = DemoStop
                        'Exit Sub
                    Else

                        If State <> "" And key_for_me <> CreateKey + "10111982" Then ' play in registry hack
                            MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì „·› «·Õ„«Ì… « ’· »„œÌ— «·‰Ÿ«„ ·Õ· «·„‘ﬂ·… Ê «·Õ’Ê· ⁄·Ï «·„› «Õ «·’ÕÌÕ", vbCritical
                            'FrmRegisteration.Show
                            'FrmRegisteration.LblExpireCount = run_count
                            SystemOptions.SysRegisterState = DemoRun
                            'Exit Sub

                        End If

                    End If
                End If
            End If
        End If
    End If

ll:
    'WriteInLogFile "Goto Check Expire" 'salimdemo
    'SystemOptions.SysRegisterState = CheckExpireation
    SystemOptions.SysRegisterState = Registered
    'SystemOptions.SysRegisterState = DemoRun
    'SystemOptions.SysRegisterState = DevelopVersion
    'SystemOptions.UserInterface = EnglishInterface

    SystemOptions.SysRunNumber = RecRead.CurRumNumber

    If SystemOptions.SysRegisterState = UnErrorOccured Then
        Msg = "⁄›Ê« ÕœÀ Œÿ« «À‰«¡ «·ﬂ‘› ⁄‰ Õ„«Ì… «·»—‰«„Ã.."
        Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï ··»—‰«„Ã"
        Msg = Msg & CHR(13) & Err.Description
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        FrmActivation.show
        '        End
    ElseIf SystemOptions.SysRegisterState = DemoStop Then
        FrmActivation.show
        '     Load FrmRegisteration
        '     FrmRegisteration.WzrdMain.CancelEnabled = False
        '     FrmRegisteration.LblExpireCount = 0
        '     FrmRegisteration.show vbModal
        '
        '        If FrmRegisteration.UserCancelReg = True Then
        '            End
        '        End If

    ElseIf SystemOptions.SysRegisterState = DemoRun Then
        '  Load FrmRegisteration
        '  FrmRegisteration.WzrdMain.CancelEnabled = True
        '
        '        FrmRegisteration.show vbModal
        '        Unload FrmRegisteration
        FrmActivation.show
    End If

    If SystemOptions.SysRegisterState = Registered Then
    
    End If
retryConection:
    If open_my_connection = False Then
        End
    End If

    If LoadMainSystemOptions = False Then
        End
    End If

    'UpdateDataBase
    '
               
    If 1 = 1 Then
        ' CheckSerial

        'Load FrmSplash
        'FrmSplash.Show
        DoEvents
        DoEvents
        
        Unload FrmLogIn

        
        
        FrmLogIn.show vbModal
        '  PutFormOnTop FrmSplash.hWnd

    Else
 
        '   Load FrmSplash
        '   FrmSplash.Show
        'PutFormOnTop FrmSplash.hwnd
    
        DoEvents
        DoEvents
        user_name = "Admin"
        user_id = 1
        User_Password = "1"
        SystemOptions.usertype = UserNourCo
        SystemOptions.UserInvoiceChangePrice = 1
        SystemOptions.UserInvoiceChangePrice1 = 1
        SystemOptions.UserInvoiceChangePrice2 = 1
            
        SystemOptions.UserInvoiceShowProfit = 1
    End If
   
    DoEvents
    
    If allowloadmdifrmmain = True Then
        Load mdifrmmain
        
        If SystemOptions.CanUploadZakatOpt And SystemOptions.CanUploadZakat Then
            Dim Frm As New FrmAnalysItems
            
           
            'FrmAnalysItems.WindowState = 1
            'FrmAnalysItems.show 1
          '  mdifrmmain.Hide
         '   mdifrmmain.Enabled = False
        End If
    Else
        End

    End If

    DoEvents
    If SystemOptions.CanUploadZakatOpt And SystemOptions.CanUploadZakat Then
       ' mdifrmmain.Hide
       ' mdifrmmain.Enabled = False
        HideAllAlarms = True
        mdifrmmain.MnuAccounts.Enabled = False
        'mdifrmmain.Help.Enabled = False
         mdifrmmain.MarketingMnu.Enabled = False
          mdifrmmain.Tools.Enabled = False
           mdifrmmain.Basicdata.Enabled = False
           mdifrmmain.tech.Enabled = False
           mdifrmmain.LIFEINDICATORMNU.Enabled = False
           'mdifrmmain.Basicdata.Enabled = False
        'mdifrmmain.BasicDataM.Enabled = False
        mdifrmmain.MdiContextMenu.Enabled = False
        
        
         Frm.mIndex = 3
            Load Frm
            
    Else
        mdifrmmain.show
    End If

    'Splish.show

    'BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "ShowRequest", True)

    'If BolShowRequest = True Then
    '    If checkApility("FrmReturnpurchases", False) = True Then
    '         If ShowRequest = True Then
    '            FrmRequest.Show
    '            FrmRequest.ZOrder 0
    '        End If
    '    End If
    'End If

    'BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "ShowPayment", True)

    'If BolShowRequest = True Then
    '    If checkApility("FrmPaymentTime", False) = True Then
    '        If ShowCurrencyAlarm = True Then
    '            FrmPaymentTime.Show
    '            FrmPaymentTime.ZOrder 0
    '        End If
    '    End If
    'End If
    'WriteTaskPanlData
    ' ⁄—÷ «·√ﬁ”«ÿ «· Ì Õ«‰ Êﬁ  ”œ«œÂ«›Ì »œ«Ì… «· Õ„Ì·
    'BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)
    'If BolShowRequest = True Then
    '    If checkApility("FrmInstallmentMustPay", False) = True Then
    '        If ShowInstallmentMustPay = True Then
    '            FrmInstallmentMustPay.Show
    '            FrmInstallmentMustPay.ZOrder 0
    '        End If
    '    End If
    'End If
    DoEvents
    DoEvents
    'Unload FrmSplash
    '''EmptyDataBase
    '⁄—÷ ‰«›–… «· ·„ÌÕ «·ÌÊ„Ì
    AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowToolTip", True)

    If AskOption = True Then
        'FrmDailyToolTip.show
    End If

    '   'alram_frm.Show 'salim2
    '  FrmDailyToolTip.Show
    'all_alarms.Show

    'FrmEmpExpir1.Show
    ' FrmEmpExpir2.Show
    'End If
    '⁄—÷  ‰»ÌÂ« ‘∆Ê‰ «·„ÊŸ›Ì‰
    'AskOption = GetSetting(StrAppRegPath, "Setting", "showhr", True)
    'If AskOption = True Then
    'all_alarms.Show
    'End If

    WebForm.show
    If mdifrmmain.CarMaintenance.Visible = True Then
        '  frmMainCars.show
    End If
        

    If HideAllAlarms = False Then

        If checkApility("System_alarms", False) = False Then
            GoTo xll
        End If

        System_alarms.show
       
        AlarmsDates
       
    End If

    If SystemOptions.OpenAtProduction = True Then
            
        FrmInstallmentVendorAlarm.show
        FrmInstallmentVendorAlarm.TabMain.CurrTab = 2
        If SystemOptions.UserInterface = ArabicInterface Then
            FrmInstallmentVendorAlarm.EleHeader.Caption = " ‰»ÌÂ«  «·«‰ «Ã"
        Else
            FrmInstallmentVendorAlarm.EleHeader.Caption = "Production Alarms"
        End If
        FrmInstallmentVendorAlarm.Caption = FrmInstallmentVendorAlarm.EleHeader.Caption
    End If
    If SystemOptions.OPenShortInvoice = True Then
        HideAllAlarms = True
        System_alarms.Hide
  
        frmsalebill5.show
  
    End If
    
    
    If SystemOptions.OPenShortInvoicePump = True Then
        HideAllAlarms = True
        System_alarms.Hide
        frmsalebill6.mTypeInvoice = 2
        frmsalebill6.show
  
    End If
       
    If SystemOptions.OPenShortInvoicePetrol = True Then
        HideAllAlarms = True
        System_alarms.Hide
        frmsalebill6.mTypeInvoice = 1
        frmsalebill6.show
  
    End If

    
 
xll:
    If GET_DEFAULT_CURRENCY_INF = False Then
        FRMcurrency.show
 
    End If

    '
 
    If ChangePW = True Then
        If SystemOptions.CanUploadZakatOpt Then
        Else
            FrmEditPW1.show
        End If
    End If
    If Messnger = True Then mdifrmmain.Timer1.Enabled = True: FrmMessnger.show

    CurrentVersion = "V25-03-2026"  'lastlast 'lastlast 'lastlast 'lastlast
    
    If CurrentVersion <> getLastDataBaseUpdateDate Then
        If SystemOptions.UserInterface = ArabicInterface Then
                                
            MsgBox "·«»œ „‰  ÕœÌÀ «·‰”Œ…", vbCritical
        Else
            MsgBox "version Must Be Updated", vbCritical
        End If
                    
        '  End
        
        AdminLogin.show
         
    End If
        
        
    If SystemOptions.OPenShortInvoice = True Then
        HideAllAlarms = True
        System_alarms.Hide
  
        frmsalebill5.show
  
    End If
 
 
     If SystemOptions.OPenShortInvoicePump = True Then
        HideAllAlarms = True
        System_alarms.Hide
        frmsalebill6.mTypeInvoice = 2
        frmsalebill6.show
  
    End If
       
    If SystemOptions.OPenShortInvoicePetrol = True Then
        HideAllAlarms = True
        System_alarms.Hide
        frmsalebill6.mTypeInvoice = 1
        frmsalebill6.show
  
    End If


    SystemOptions.HaveTaxOnSalles = False
    '############################################# 'salim3
    'Dim RsOptions As New ADODB.Recordset
    '        Set RsOptions = New ADODB.Recordset
    '        RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
 
    'OpenScreen InvoiceScreen
    '
    '#########################################################
      Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    'XXXXXX

    Msg = "⁄›Ê« ÕœÀ  „‘ﬂ·… √À‰«¡  Õ„Ì· «·»—‰«„Ã "
    Msg = Msg & CHR(13) & "Err.Description" & Err.Description
    Msg = Msg & CHR(13) & "Err.Number:" & Err.Number
    Msg = Msg & CHR(13) & "Err.Source:" & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '     save_login_info1 "Byte", "«·ﬁ«⁄œ… «·«”«”Ì…"
    'Load FrmSQLConData
    '   FrmSQLConData.show vbModal
    open_my_connection True
    GoTo retryConection
    
    'Dim RsOption As New ADODB.Recordset
    'RsOption.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    'RsOption("RunCount").Value = 0
    'RsOption.update
    'RsOption.Close

    'SystemOptions.SysRegsAppPath = App.Path & "\Barcode"
    'SystemOptions.SysVersion = RegisterVersion

End Sub

Public Function GetMsgs(IntCode As Integer, _
                        IntButtons As VBA.VbMsgBoxStyle) As VBA.VbMsgBoxResult
    Dim Msg     As String
    Dim BoxRtl  As Long
    Dim StrPath As String
    Dim IntFile As Integer
    Dim MSGType As MsgData
    StrPath = App.path
    StrPath = IIf(right(StrPath, 1) = "\", "", StrPath & "\")
    StrPath = StrPath & "Msgs.dat"

    If Dir(StrPath, vbNormal) = "" Then
        'CreateMsgsFile
    End If

    'Msg = "Sorry.!" & Chr(13)
    'Msg = "The saving is failed"
    'Msg = "This Record Is Saved." & Chr(13)
    'Msg = Msg & "Do You Want To Open New Record.?"
    IntFile = FreeFile
    Open StrPath For Random As #IntFile Len = Len(MSGType)
    Get #IntFile, IntCode, MSGType
    Close #IntFile

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = Trim(MSGType.MsgArabic)
        BoxRtl = VBA.VbMsgBoxStyle.vbMsgBoxRight + VBA.VbMsgBoxStyle.vbMsgBoxRtlReading
    Else
        Msg = Trim(MSGType.MsgEnglish)
        BoxRtl = 0
    End If

    GetMsgs = MsgBox(Msg, IntButtons + BoxRtl, App.Title)
End Function

Public Function updateversion(Version As String, funId As Integer)
    Dim sql As String
    Version = "" & Version & ""
    sql = "Update Systemversion Set version='" & Version & "' ,funId=" & funId & " where id=1"
    Cn.Execute sql
End Function

Function updateProcedure3()
    On Error Resume Next
 
    Dim sql As String

    sql = "    DROP FUNCTION QryItemsSalesTotal" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION QryItemsSalesTotal(@TransType int =0,@TransType2 int=0,@TransType3 int=0,@FromDate datetime ,@ToDate datetime,@ItemType int=0 )" & CHR(13)
    sql = sql & " RETURNS @xTable TABLE" & CHR(13)
    sql = sql & "   (" & CHR(13)
    sql = sql & " ItemID int," & CHR(13)
    sql = sql & " ItemCode nvarchar(50)," & CHR(13)
    sql = sql & " ItemName nvarchar(255)," & CHR(13)
    sql = sql & " GroupID  int," & CHR(13)
    sql = sql & " Total   money," & CHR(13)
    sql = sql & "     totalqty Float" & CHR(13)
    sql = sql & "    )" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)

    sql = sql & " INSERT @xTable" & CHR(13)
    sql = sql & " Select ItemID,ItemCode,ItemName,GroupID,Sum(Total) as Totals,Sum(Quantity) as TotalQty" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & "     SELECT TblItems.ItemID,TblItems.ItemCode, TblItems.ItemName,TblItems.GroupID," & CHR(13)
    sql = sql & " 'Total'=Case" & CHR(13)
    sql = sql & " When ItemDiscountType=1 Or ItemDiscountType=0 Then (Transaction_Details.SHOWQTY*Transaction_Details.SHOWPRICE-isnull(Transaction_Details.TotalDiscountPerLine,0))*isnull(Currency_rate,1)" & CHR(13)
    sql = sql & " When ItemDiscountType=2 Then ( ((Transaction_Details.SHOWQTY*Transaction_Details.SHOWPRICE)-ItemDiscount)-isnull(Transaction_Details.TotalDiscountPerLine,0))*isnull(Currency_rate,1)" & CHR(13)
    sql = sql & " When ItemDiscountType=3 Then ( (Transaction_Details.SHOWQTY*Transaction_Details.SHOWPRICE) *( 1- (ItemDiscount/100))-isnull(Transaction_Details.TotalDiscountPerLine,0))*isnull(Currency_rate,1)" & CHR(13)
    sql = sql & " Else  0" & CHR(13)
    sql = sql & " End" & CHR(13)
    sql = sql & " ,Transaction_Details.Quantity" & CHR(13)
    sql = sql & " FROM dbo.TblItems INNER JOIN  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN" & CHR(13)
    sql = sql & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID" & CHR(13)
    sql = sql & " WHERE (Transactions.Transaction_Type=@TransType  OR Transactions.Transaction_Type=@TransType2 OR Transactions.Transaction_Type=@TransType3 )" & CHR(13)
                         
    sql = sql & "     AND" & CHR(13)
    sql = sql & "     Transactions.Transaction_Date >=@FromDate" & CHR(13)
    sql = sql & " AND" & CHR(13)
    sql = sql & " Transactions.Transaction_Date <=@TODate" & CHR(13)
    sql = sql & " and dbo.TblItems.ItemType=@ItemType" & CHR(13)
    sql = sql & " )DrivTable" & CHR(13)
    sql = sql & " Group By ItemID,ItemCode,ItemName,GroupID" & CHR(13)
    sql = sql & " Return" & CHR(13)
    sql = sql & " End" & CHR(13)
    db_createOrUpdateFuctionSQL "QryItemsSalesTotal", sql
End Function

Function updateProcedure2()
    On Error Resume Next
 
    Dim sql As String

    sql = "    DROP FUNCTION GetOpeningBalance" & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION GetOpeningBalance(@fromdate datetime ,@accountcode as varchar(255) ,@LastAccount as integer)" & CHR(13)
    sql = sql & " RETURNS Float" & CHR(13)
    sql = sql & "  AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " RETURN (" & CHR(13)
    sql = sql & " SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & " FROM (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
                
    sql = sql & " From dbo.DOUBLE_ENTREY_VOUCHERS1" & CHR(13)
    'sql = sql & " where (dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like @accountcode + '%' )"
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDate  =@fromdate" & CHR(13)
              
    sql = sql & " )XTABLE" & CHR(13)
    sql = sql & " )" & CHR(13)

    sql = sql & " End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetOpeningBalance", sql

    sql = "    DROP FUNCTION GetOpeningBalanceByActivity" & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION GetOpeningBalanceByActivity(@fromdate datetime  ,@Activity_Id as integer,@accountcode as varchar(255), @LastAccount as integer )" & CHR(13)
    sql = sql & "  RETURNS Float" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " RETURN (" & CHR(13)
    sql = sql & " SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & " FROM (" & CHR(13)
    sql = sql & "              SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & "   DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
                
    sql = sql & " from dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN" & CHR(13)
    sql = sql & " dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
    'sql = sql & " where (dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like @accountcode + '%' ) "
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDate  =@fromdate and dbo.TblBranchesData.ActivityTypeId =@Activity_Id" & CHR(13)
              
    sql = sql & " )XTABLE" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "GetOpeningBalanceByActivity", sql
 
    sql = "    DROP FUNCTION GetOpeningBalanceByBranch" & CHR(13)
    Cn.Execute sql

    sql = "CREATE FUNCTION GetOpeningBalanceByBranch(@fromdate datetime  ,@Branch_Id as integer,@accountcode as varchar(255),@LastAccount as integer )" & CHR(13)
    sql = sql & "  RETURNS Float" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " RETURN (" & CHR(13)
    sql = sql & " SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & " FROM (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & "  When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
                
    sql = sql & " from dbo.DOUBLE_ENTREY_VOUCHERS1 INNER JOIN" & CHR(13)
    sql = sql & " dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS1.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
    'sql = sql & " where (dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like @accountcode + '%' ) "
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS1.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS1.RecordDate  =@fromdate and dbo.DOUBLE_ENTREY_VOUCHERS1.Branch_Id =@Branch_Id" & CHR(13)
              
    sql = sql & " )XTABLE" & CHR(13)
    sql = sql & " )" & CHR(13)

    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "GetOpeningBalanceByBranch", sql
 
    '$$$$$$$$$$$$$$$$$part2
    sql = "    DROP FUNCTION GetBalance" & CHR(13)
    Cn.Execute sql
 
    sql = "CREATE FUNCTION GetBalance(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@LastAccount as integer)" & CHR(13)
    sql = sql & "  RETURNS Float" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " RETURN (" & CHR(13)
    sql = sql & " SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & " FROM (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
                
    sql = sql & " From dbo.DOUBLE_ENTREY_VOUCHERS" & CHR(13)
    'sql = sql & " where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' ) " & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & "and  ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
              
    sql = sql & " )XTABLE" & CHR(13)
    sql = sql & " )" & CHR(13)

    sql = sql & "  End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "GetBalance", sql
 
    sql = "    DROP FUNCTION GetBalanceByActivity" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION GetBalanceByActivity(@fromdate datetime,@Todate datetime  ,@Activity_Id as integer,@accountcode as varchar(255),@LastAccount as integer)" & CHR(13)
    sql = sql & "   RETURNS Float" & CHR(13)
    sql = sql & "   AS" & CHR(13)
    sql = sql & "  Begin" & CHR(13)
    sql = sql & "  RETURN (" & CHR(13)
    sql = sql & "  SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & "  FROM (" & CHR(13)
    sql = sql & "   SELECT" & CHR(13)
    sql = sql & "  DEV_Value1=Case" & CHR(13)
    sql = sql & "  When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & "  END," & CHR(13)
    sql = sql & "  DEV_Value2=Case" & CHR(13)
    sql = sql & "  When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & "  End" & CHR(13)
                
    sql = sql & "  from dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN" & CHR(13)
    sql = sql & "   dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
    'sql = sql & "  where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' )" & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & "   and     ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
    sql = sql & "    and dbo.TblBranchesData.ActivityTypeId =@Activity_Id" & CHR(13)
              
    sql = sql & "  )XTABLE" & CHR(13)
    sql = sql & "  )" & CHR(13)

    sql = sql & "  End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetBalanceByActivity", sql
   
    sql = "    DROP FUNCTION GetBalanceByBranch" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION GetBalanceByBranch(@fromdate datetime  ,@Todate datetime  ,@Branch_Id as integer,@accountcode as varchar(255),@LastAccount as integer)" & CHR(13)
    sql = sql & "     RETURNS Float" & CHR(13)
    sql = sql & "    AS" & CHR(13)
    sql = sql & "    Begin" & CHR(13)
    sql = sql & "    RETURN (" & CHR(13)
    sql = sql & "    SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & "    FROM (" & CHR(13)
    sql = sql & "    SELECT" & CHR(13)
    sql = sql & "    DEV_Value1=Case" & CHR(13)
    sql = sql & "    When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "    Else 0" & CHR(13)
    sql = sql & "    END," & CHR(13)
    sql = sql & "    DEV_Value2=Case" & CHR(13)
    sql = sql & "    When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "    Else 0" & CHR(13)
    sql = sql & "    End" & CHR(13)
                
    sql = sql & "    from dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN" & CHR(13)
    sql = sql & "     dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
    ' sql = sql & "    where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' ) and" & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)

    sql = sql & "   and  ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
    sql = sql & "    and dbo.DOUBLE_ENTREY_VOUCHERS.Branch_Id =@Branch_Id" & CHR(13)
              
    sql = sql & "    )XTABLE" & CHR(13)
    sql = sql & "    )" & CHR(13)

    sql = sql & "     End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "GetBalanceByBranch", sql
   
    sql = "    DROP FUNCTION GetBalanceCreditORdepit" & CHR(13)
    Cn.Execute sql
   
    sql = "CREATE FUNCTION GetBalanceCreditORdepit(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer ,@LastAccount as integer)" & CHR(13)
    sql = sql & "      RETURNS Float" & CHR(13)
    sql = sql & "     AS" & CHR(13)
    sql = sql & "     Begin" & CHR(13)
    sql = sql & "      RETURN (" & CHR(13)
    sql = sql & "     SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & "     FROM (" & CHR(13)
    sql = sql & "     SELECT" & CHR(13)
    sql = sql & "     DEV_Value1=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     END," & CHR(13)
    sql = sql & "     DEV_Value2=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     End" & CHR(13)
    sql = sql & "     From dbo.DOUBLE_ENTREY_VOUCHERS" & CHR(13)
    'sql = sql & "     where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' ) and" & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & "  and   ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
    sql = sql & "     and(Credit_Or_Debit=@Credit_Or_Debit)" & CHR(13)
 
    sql = sql & "      )XTABLE" & CHR(13)
    sql = sql & "     )" & CHR(13)
    sql = sql & "     End" & CHR(13)

    db_createOrUpdateFuctionSQL "GetBalanceCreditORdepit", sql

    sql = "    DROP FUNCTION GetBalanceCreditORdepitByActivity" & CHR(13)
    Cn.Execute sql
  
    sql = "CREATE FUNCTION GetBalanceCreditORdepitByActivity(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer,@Activity_Id as integer ,@LastAccount as integer)" & CHR(13)
    sql = sql & "       RETURNS Float" & CHR(13)
    sql = sql & "     AS" & CHR(13)
    sql = sql & "     Begin" & CHR(13)
    sql = sql & "     RETURN (" & CHR(13)
    sql = sql & "     SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & "     FROM (" & CHR(13)
    sql = sql & "     SELECT" & CHR(13)
    sql = sql & "     DEV_Value1=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     END," & CHR(13)
    sql = sql & "      DEV_Value2=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     End" & CHR(13)
    sql = sql & "     from dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN" & CHR(13)
    sql = sql & "     dbo.TblBranchesData ON dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = dbo.TblBranchesData.branch_id" & CHR(13)
    'sql = sql & "     where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' )" & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & "     and     ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
    sql = sql & "     and dbo.TblBranchesData.ActivityTypeId =@Activity_Id" & CHR(13)

    sql = sql & "     and(Credit_Or_Debit=@Credit_Or_Debit)" & CHR(13)

    sql = sql & "     )XTABLE" & CHR(13)
    sql = sql & "     )" & CHR(13)
    sql = sql & "     End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetBalanceCreditORdepitByActivity", sql

    sql = "    DROP FUNCTION GetBalanceCreditORdepitByBranch" & CHR(13)
    Cn.Execute sql
   
    sql = "CREATE FUNCTION GetBalanceCreditORdepitByBranch(@fromdate datetime,@Todate datetime ,@accountcode as varchar(255),@Credit_Or_Debit as integer ,@Branch_Id as integer,@LastAccount as integer)" & CHR(13)
    sql = sql & "      RETURNS Float" & CHR(13)
    sql = sql & "     AS" & CHR(13)
    sql = sql & "     Begin" & CHR(13)
    sql = sql & "      RETURN (" & CHR(13)
    sql = sql & "     SELECT     Sum(DEV_Value1)-Sum(DEV_Value2) as  result" & CHR(13)
    sql = sql & "     FROM (" & CHR(13)
    sql = sql & "     SELECT" & CHR(13)
    sql = sql & "     DEV_Value1=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=0   Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     END," & CHR(13)
    sql = sql & "     DEV_Value2=Case" & CHR(13)
    sql = sql & "     When Credit_Or_Debit=1  Then Value * 1" & CHR(13)
    sql = sql & "     Else 0" & CHR(13)
    sql = sql & "     End" & CHR(13)
    sql = sql & "     From dbo.DOUBLE_ENTREY_VOUCHERS" & CHR(13)
    ' sql = sql & "     where (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like @accountcode + '%' ) and" & Chr(13)
    sql = sql & " WHERE dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code like CASE WHEN @LastAccount=1 THEN @accountcode ELSE  @accountcode + 'a%'  END" & CHR(13)
    sql = sql & "  and   ( dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  >=@fromdate   and dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate  <=@Todate  )" & CHR(13)
    sql = sql & "     and(Credit_Or_Debit=@Credit_Or_Debit)" & CHR(13)
    sql = sql & "       and dbo.DOUBLE_ENTREY_VOUCHERS.Branch_Id =@Branch_Id" & CHR(13)
    sql = sql & "      )XTABLE" & CHR(13)
    sql = sql & "     )" & CHR(13)
    sql = sql & "     End" & CHR(13)

    db_createOrUpdateFuctionSQL "GetBalanceCreditORdepitByBranch", sql

    sql = "    DROP FUNCTION GetItemCostPrice" & CHR(13)
    Cn.Execute sql
    sql = "CREATE FUNCTION GetItemCostPrice(@fromdate datetime,@Todate datetime ,@itemid as integer)" & CHR(13)
    sql = sql & "      RETURNS Float" & CHR(13)
    sql = sql & "     AS" & CHR(13)
    sql = sql & "    Begin" & CHR(13)
    sql = sql & "    RETURN (select  round( Total / TotalQty, 5) AS AvCost  from dbo.QryItemsTransactionsTotals(28, 3,20, @fromdate, @Todate) where itemid=@itemid)" & CHR(13)
    sql = sql & "    End" & CHR(13)
    db_createOrUpdateFuctionSQL "GetItemCostPrice", sql

    db_createOrUpdateFuctionSQL "GetItemUnitFactor", sql

    sql = " CREATE FUNCTION GetItemUnitFactor(@itemid as integer,@unitID as integer)"
    sql = sql & "      RETURNS Float"
    sql = sql & "      AS"
    sql = sql & "      Begin"
    sql = sql & "      RETURN (select   UnitFactor   from   dbo.TblItemsUnits   WHERE     (UnitID = @unitID) AND (ItemID = @itemid ))"
    sql = sql & "      End"
    db_createOrUpdateFuctionSQL "GetItemUnitFactor", sql

End Function

Function updateProcedure()
    On Error Resume Next
 
    Dim sql As String

    sql = sql & "    DROP FUNCTION QryIncomeStatement1" & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryIncomeStatement1] (@fromdate datetime,@todate datetime )" & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account int," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " AccountTypes int" & CHR(13)
    
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & "  Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account,AccountTypes" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on  dbo.RptLedger_Sub.Account_Code=dbo.Accounts.Account_Code" & CHR(13)
 
    sql = sql & " where dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate" & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes=2 )" & CHR(13)
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)
    sql = sql & "         Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryIncomeStatement1", sql

    sql = ""

    sql = sql & "    DROP FUNCTION QryIncomeStatementActivity " & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryIncomeStatementActivity] (@fromdate datetime,@todate datetime,@ActivityTypeId as integer)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account int," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " AccountTypes int" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & "  Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account,AccountTypes" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on  dbo.RptLedger_Sub.Account_Code=dbo.Accounts.Account_Code" & CHR(13)
 
    sql = sql & " where dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate and dbo.RptLedger_Sub.ActivityTypeId=@ActivityTypeId " & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes=2 )" & CHR(13)
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)
    sql = sql & "         Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryIncomeStatementActivity", sql

    sql = ""

    sql = sql & "    DROP FUNCTION QryIncomeStatementBranch " & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryIncomeStatementBranch] (@fromdate datetime,@todate datetime,@Branch_Id as integer)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account int," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " AccountTypes int" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & "  Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account,AccountTypes" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on  dbo.RptLedger_Sub.Account_Code=dbo.Accounts.Account_Code" & CHR(13)
 
    sql = sql & " where dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate and dbo.RptLedger_Sub.Branch_Id =@Branch_Id  " & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes=2 )" & CHR(13)
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,AccountTypes" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)
    sql = sql & "         Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryIncomeStatementBranch", sql

    sql = "    DROP FUNCTION QryBalanceSheet" & CHR(13)

    Cn.Execute sql
    sql = ""
    sql = "CREATE FUNCTION [dbo].[QryBalanceSheet] (@fromdate datetime,@todate datetime)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account int," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)" & CHR(13)
              
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & " Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code" & CHR(13)
    sql = sql & " from" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account, AccountTypes" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on" & CHR(13)
    sql = sql & " dbo.RptLedger_Sub.Account_Code = dbo.Accounts.Account_Code" & CHR(13)
    sql = sql & " where (dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate)  or (   dbo.RptLedger_Sub.RecordDate is null)" & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes = 1)" & CHR(13)
 
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)

    sql = sql & " Return" & CHR(13)
    sql = sql & " End" & CHR(13)
  
    db_createOrUpdateFuctionSQL "QryBalanceSheet", sql
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    sql = ""

    sql = sql & "    DROP FUNCTION QryBalanceSheetActivity " & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryBalanceSheetActivity] (@fromdate datetime,@todate datetime,@ActivityTypeId as integer)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account bit," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " AccountTypes int," & CHR(13)
    sql = sql & "  opening_balance Decimal(18, 2)," & CHR(13)
    sql = sql & "  balance Decimal(18, 2)" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & "  Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,AccountTypes,opening_balance,balance" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account,AccountTypes,dbo.Accounts.opening_balance,balance" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on  dbo.RptLedger_Sub.Account_Code=dbo.Accounts.Account_Code" & CHR(13)
 
    sql = sql & " where dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate and dbo.RptLedger_Sub.ActivityTypeId=@ActivityTypeId " & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes=1 )" & CHR(13)
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,AccountTypes,opening_balance,balance" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)
    sql = sql & "         Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryBalanceSheetActivity", sql

    sql = ""

    sql = sql & "    DROP FUNCTION QryBalanceSheetBranch " & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryBalanceSheetBranch] (@fromdate datetime,@todate datetime,@Branch_Id as integer)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & " Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account bit," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " AccountTypes int," & CHR(13)
    sql = sql & "  opening_balance Decimal(18, 2)," & CHR(13)
    sql = sql & "  balance Decimal(18, 2)" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & "  Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & " account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,AccountTypes,opening_balance,balance" & CHR(13)
    sql = sql & " From" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account,AccountTypes,dbo.Accounts.opening_balance,balance" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on  dbo.RptLedger_Sub.Account_Code=dbo.Accounts.Account_Code" & CHR(13)
 
    sql = sql & " where dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate and dbo.RptLedger_Sub.Branch_Id =@Branch_Id  " & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes=1 )" & CHR(13)
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,AccountTypes,opening_balance,balance" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)
    sql = sql & "         Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryBalanceSheetBranch", sql
    'balanceSheet Normal

    sql = ""

    sql = sql & "    DROP FUNCTION QryBalanceSheet " & CHR(13)

    Cn.Execute sql
    sql = "CREATE FUNCTION [dbo].[QryBalanceSheet] (@fromdate datetime,@todate datetime)  " & CHR(13)
    sql = sql & " RETURNS @XTable Table" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " DebitValue Decimal (18,2)," & CHR(13)
    sql = sql & " CreditValue Decimal (18,2)," & CHR(13)
    sql = sql & " Account_Code nvarchar(255)," & CHR(13)
    sql = sql & " Account_Name nvarchar(255)," & CHR(13)
    sql = sql & " Account_Serial nvarchar(255)," & CHR(13)
    sql = sql & "   Account_NameEng nvarchar(255)," & CHR(13)
    sql = sql & " last_Account bit," & CHR(13)
    sql = sql & " Parent_Account_Code nvarchar(255)," & CHR(13)
    sql = sql & "  opening_balance Decimal(18, 2)," & CHR(13)
    sql = sql & "  balance Decimal(18, 2)" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)

    sql = sql & " Begin" & CHR(13)
    sql = sql & " INSERT  @XTable" & CHR(13)
    sql = sql & " Select Sum(DEV_Value1) as DebitValue,Sum(DEV_Value2) as CreditValue,Account_Code," & CHR(13)
    sql = sql & "   account_name , account_serial, Account_NameEng, last_account, Parent_Account_Code,opening_balance,balance" & CHR(13)
    sql = sql & " from" & CHR(13)
    sql = sql & "  (" & CHR(13)
    sql = sql & " SELECT" & CHR(13)
    sql = sql & " DEV_Value1=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=0   Then DEV_Value * 1" & CHR(13)
    sql = sql & "  Else 0" & CHR(13)
    sql = sql & " END," & CHR(13)
    sql = sql & " DEV_Value2=Case" & CHR(13)
    sql = sql & " When Credit_Or_Debit=1  Then DEV_Value * 1" & CHR(13)
    sql = sql & " Else 0" & CHR(13)
    sql = sql & " End" & CHR(13)
    
    sql = sql & " , dbo.Accounts.Account_Code, dbo.Accounts.Account_Name, dbo.Accounts.Account_Serial," & CHR(13)
    sql = sql & " dbo.Accounts.Account_NameEng , dbo.Accounts.Parent_Account_Code, dbo.Accounts.last_account, AccountTypes,dbo.Accounts.opening_balance,balance" & CHR(13)
    sql = sql & " FROM         dbo.RptLedger_Sub Right Join dbo.Accounts on" & CHR(13)
    sql = sql & " dbo.RptLedger_Sub.Account_Code = dbo.Accounts.Account_Code" & CHR(13)
    sql = sql & " where (dbo.RptLedger_Sub.RecordDate >=@fromdate and  dbo.RptLedger_Sub.RecordDate <= @todate)  or (   dbo.RptLedger_Sub.RecordDate is null)" & CHR(13)
    sql = sql & " )XTable" & CHR(13)
    sql = sql & " Where (AccountTypes = 1)" & CHR(13)
 
    sql = sql & " Group by Account_Code, Account_Name, Account_Serial, Account_NameEng,last_Account,Parent_Account_Code,opening_balance,balance" & CHR(13)
    sql = sql & " Order by Account_Code ASC" & CHR(13)

    sql = sql & " Return" & CHR(13)
    sql = sql & " End" & CHR(13)
 
    db_createOrUpdateFuctionSQL "QryBalanceSheet", sql

    'balanceSheet Normal

    sql = ""

    sql = sql & "    DROP FUNCTION QryTransactionsTotal " & CHR(13)

    Cn.Execute sql

    sql = "CREATE FUNCTION [dbo].[QryTransactionsTotal] (  )" & CHR(13)
    sql = sql & " RETURNS @xTable TABLE" & CHR(13)
    sql = sql & " (" & CHR(13)
    
    sql = sql & " Transaction_ID  int," & CHR(13)
    sql = sql & " TransSum float," & CHR(13)
    sql = sql & " TransNet float," & CHR(13)
    sql = sql & " Transaction_Serial nvarchar(50)," & CHR(13)
    sql = sql & " Transaction_Date smalldatetime," & CHR(13)
    sql = sql & " Transaction_Type int," & CHR(13)
    sql = sql & " PaymentType int," & CHR(13)
    sql = sql & " Transaction_HijriDate   smalldatetime," & CHR(13)
    sql = sql & " Trans_Discount  real," & CHR(13)
    sql = sql & " Trans_DiscountType  int," & CHR(13)
    sql = sql & " CusID   int," & CHR(13)
    sql = sql & " StoreID int," & CHR(13)
    sql = sql & " UserID  int," & CHR(13)
    sql = sql & " Emp_ID  int," & CHR(13)
    sql = sql & " TaxFound bit," & CHR(13)
    sql = sql & " TaxValue real," & CHR(13)
    sql = sql & " TransProfit money," & CHR(13)
    sql = sql & " TotalAfterTax float," & CHR(13)
    sql = sql & " ReturnID int" & CHR(13)
    sql = sql & " )" & CHR(13)
    sql = sql & " AS" & CHR(13)
    sql = sql & " Begin" & CHR(13)
    sql = sql & " INSERT @xTable" & CHR(13)
    sql = sql & " SELECT Transaction_ID,TransSum,TransNet" & CHR(13)
    sql = sql & " , Transaction_Serial,Transaction_Date,Transaction_Type,PaymentType,Transaction_HijriDate,Trans_Discount," & CHR(13)
    sql = sql & " Trans_DiscountType,CusID,StoreID,UserID,Emp_ID,TaxFound,TaxValue, ItemProfit," & CHR(13)
    sql = sql & " 'TotalAfterTax'=Case" & CHR(13)
    sql = sql & " WHEN TaxFound =1 THEN Round((TransNet) * (1+ (TaxValue/100)),2)" & CHR(13)
    sql = sql & " Else" & CHR(13)
    sql = sql & " Round((TransNet),2)" & CHR(13)
    sql = sql & " End" & CHR(13)
    sql = sql & " ,ReturnID" & CHR(13)
    sql = sql & " from" & CHR(13)
    sql = sql & " (" & CHR(13)

    sql = sql & " SELECT Transaction_ID," & CHR(13)
    sql = sql & " [TransSum]=Round(Sum(TransSum),2)," & CHR(13)
    sql = sql & "   [TransNet]=CASE" & CHR(13)
    sql = sql & " WHEN Trans_DiscountType=0 THEN Round(Sum(TransSum),2)" & CHR(13)
    sql = sql & " WHEN Trans_DiscountType=1 THEN Round((Sum(TransSum)-Trans_Discount),2)" & CHR(13)
    sql = sql & " WHEN Trans_DiscountType=2 THEN Round((Sum(TransSum) * (1-[Trans_Discount]/100)),2)" & CHR(13)
    sql = sql & " Else" & CHR(13)
    sql = sql & " Round(Sum(TransSum),2)" & CHR(13)
    sql = sql & " End" & CHR(13)
    sql = sql & " , Transaction_Serial,Transaction_Date,Transaction_Type,PaymentType,Transaction_HijriDate,Trans_Discount," & CHR(13)
    sql = sql & " Trans_DiscountType,CusID,StoreID,UserID,Emp_ID,TaxFound,TaxValue,ReturnID,Sum(ItemProfit) As ItemProfit" & CHR(13)
    sql = sql & " from" & CHR(13)
    sql = sql & " (" & CHR(13)
    sql = sql & " SELECT Transactions.Transaction_ID," & CHR(13)
    sql = sql & " 'TransSum'=Case" & CHR(13)
    sql = sql & " When ItemDiscountType=1 Or ItemDiscountType=0 Then Transaction_Details.ShowQty*Transaction_Details.showPrice" & CHR(13)
    sql = sql & " When ItemDiscountType=2 Then ((Transaction_Details.ShowQty*Transaction_Details.showPrice)-ItemDiscount)" & CHR(13)
    sql = sql & " When ItemDiscountType=3 Then (Transaction_Details.ShowQty*Transaction_Details.showPrice)-((Transaction_Details.ShowQty*Transaction_Details.showPrice)*(ItemDiscount/100))" & CHR(13)
    sql = sql & "  Else  0" & CHR(13)
            
    sql = sql & "  End ," & CHR(13)
    sql = sql & " Transactions.Transaction_Serial, Transactions.Transaction_Date, Transactions.Transaction_Type, Transactions.PaymentType," & CHR(13)
    sql = sql & " Transactions.Transaction_HijriDate, Transactions.Trans_Discount, Transactions.Trans_DiscountType, Transactions.CusID," & CHR(13)
    sql = sql & " Transactions.Storeid , Transactions.UserID, Transactions.Emp_id, Transactions.TaxFound, Transactions.TaxValue, ReturnID, Transaction_Details.ItemProfit" & CHR(13)
    sql = sql & " FROM Transactions INNER JOIN Transaction_Details ON Transactions.Transaction_ID=Transaction_Details.Transaction_ID" & CHR(13)
    sql = sql & " )YTable" & CHR(13)
    sql = sql & " Group By Transaction_ID, Transaction_Serial,Transaction_Date,Transaction_Type,PaymentType," & CHR(13)
    sql = sql & " Transaction_HijriDate , Trans_Discount, Trans_DiscountType, CusID, Storeid, UserID, Emp_id, TaxFound, TaxValue, ReturnID" & CHR(13)
    sql = sql & " )YYTable" & CHR(13)
    sql = sql & " Return" & CHR(13)
    sql = sql & " End" & CHR(13)

    db_createOrUpdateFuctionSQL "QryTransactionsTotal", sql

End Function

Private Sub CheckTblEmpPassOver()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    StrSQL = "select * form TblEmpPassOver "
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open "TblEmpPassOver", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    'RsTemp.Open Strsql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If RsTemp.RecordCount = 0 Then
        StrSQL = "DROP TABLE TblEmpPassOver"
        Cn.Execute StrSQL
                
        If DB_CreateTable("TblEmpPassOver", True, "AdvanceID", True) = True Then
            DB_CreateField "TblEmpPassOver", "Branch_NO", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "Emp_id", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "interval", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "UserID", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "AdvanceDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "DeparmentID", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "JobTypeID", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "Remark", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
            DB_CreateField "TblEmpPassOver", "Expectedouttime", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "ExpectedIntime", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "Actualouttime", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "ActualIntime", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "OutTypeID", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "PostedDate", adDBTimeStamp, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "NoteSerial", adVarWChar, adColNullable, 255, , "???   ", False, True, , True
            DB_CreateField "TblEmpPassOver", "Approved", adBoolean, adColNullable, , , "???? ?? ??", False, True
            DB_CreateField "TblEmpPassOver", "Transaction_ID", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "TblEmpPassOver", "Posted", adInteger, adColNullable, , , "      ", False, True
            DB_CreateField "TblEmpPassOver", "ComputerNo", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "Name", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "EmpType", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "DcbLeaderID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "LeaderName", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "Nationality", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "NumID", adVarWChar, adColNullable, 400, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "Remarks2", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "BoardNO", adVarWChar, adColNullable, 100, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "OperatorN", adVarWChar, adColNullable, 100, , "      ", False, True, , True
            DB_CreateField "TblEmpPassOver", "ModelID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "ColorID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "TypeEqupID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "EquepmentID", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "BossId", adInteger, adColNullable, , , " ???    ", False, True
            DB_CreateField "TblEmpPassOver", "BossNotes", adVarWChar, adColNullable, 1500, , "      ", False, True, , True

            '  FrmCheckSerial.show vbModal
        End If
                
    Else
        Exit Sub
    End If
 
ErrTrap:
End Sub

Public Sub SetInterface1(Frm As Form)
    On Error Resume Next
    Dim Txt    As TextBox
    Dim lbl    As Label
    Dim xTab   As C1Tab
    Dim Temp   As C1Tab
    Dim ctl    As Control
    Dim ConCtl As Control
    Dim i      As Integer
    'Dim dd As VSFlex8Ctl.VSFlexGrid
    Frm.RightToLeft = False

    For Each ConCtl In Frm.Controls
 
        If TypeOf ConCtl Is frame Then
            ConCtl.RightToLeft = False
            ConCtl.top = ConCtl.top

            If TypeOf ConCtl.Container Is Form Then
                ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
            Else
                ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
            End If

        ElseIf TypeOf ConCtl Is C1Elastic Then

            If Not TypeOf ConCtl.Container Is C1Tab Then
                If TypeOf ConCtl.Container Is Form Then
                    ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
                Else
                    ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
                End If
            End If

            If ConCtl.CaptionPos = cpRightCenter Then
                ConCtl.CaptionPos = cpLeftCenter
            ElseIf ConCtl.CaptionPos = cpRightTop Then
                ConCtl.CaptionPos = cpLeftTop
            End If

            If ConCtl.PicturePos = ppLeftCenter Then
                ConCtl.PicturePos = ppRightCenter
            ElseIf ConCtl.PicturePos = ppLeftTop Then
                ConCtl.PicturePos = ppRightTop
            End If

        ElseIf TypeOf ConCtl Is PictureBox Then
            ConCtl.RightToLeft = False

            If TypeOf ConCtl.Container Is Form Then
                ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
            Else
                ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
            End If
        End If

    Next ConCtl

    For Each ctl In Frm.Controls

        If Not (TypeOf ctl Is Timer Or TypeOf ctl Is CommonDialog Or TypeOf ctl Is line Or TypeOf ctl Is ImageList Or TypeOf ctl Is frame Or TypeOf ctl Is C1Elastic Or TypeOf ctl Is C1Tab Or TypeOf ctl Is Menu Or TypeOf ctl Is PictureBox Or TypeOf ctl Is XPPopUp30) Then

            If TypeOf ctl Is TextBox Then
                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is Label Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is ISAniLabel Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is DTPicker Then

                If ctl.Format = dtpCustom Then
                    ctl.CustomFormat = "d/M/yyyy"
                End If

            ElseIf TypeOf ctl Is DataCombo Then
                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is CheckBox Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is OptionButton Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is CommandButton Then
                ctl.RightToLeft = False
            ElseIf (TypeOf ctl Is ISButton) Or (TypeOf ctl Is ISButtonLW) Then
                ctl.RightToLeft = False

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                If ctl.ButtonPositionImage = impRightOfText Then
                    ctl.ButtonPositionImage = impLeftOfText
                End If

            ElseIf TypeOf ctl Is ComboBox Then
                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is Image Then
            ElseIf (TypeOf ctl Is VSFlex8Ctl.VSFlexGrid) Or (TypeOf ctl Is VSFlex8UCtl.VSFlexGrid) Then
                ctl.RightToLeft = False

                For i = 0 To ctl.Cols - 1

                    If ctl.ColAlignment(i) = flexAlignRightCenter Then
                        ctl.ColAlignment(i) = flexAlignLeftCenter
                    End If

                    If ctl.FixedAlignment(i) = flexAlignRightCenter Then
                        ctl.FixedAlignment(i) = flexAlignLeftCenter
                    End If

                Next i

                'ElseIf TypeOf Ctl Is AniGIF Then
            
            End If

            'If Ctl.Name = "CBtnColor" Then Stop
            If TypeOf ctl.Container Is Form Then
                ctl.left = ctl.Container.ScaleWidth - (ctl.left + ctl.Width)
            Else
                ctl.left = ctl.Container.Width - (ctl.left + ctl.Width)
            End If
        End If

    Next ctl

End Sub

Public Sub CreatLog_File_for_GetHeaders(str As String, _
                                        Optional FileName As String, _
                                        Optional ss As String)
    Dim StrLogFileName As String
    Dim IntFreeFile    As Integer
    'Dim ss As String

    StrLogFileName = App.path & "\Titles\" & FileName & ".txt"

    If Dir(StrLogFileName) <> "" Then
        '   Kill StrLogFileName
    End If
 
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub
 
Public Function FindIndex(ByRef F As Form, ByRef ctl As Control) As Integer
    Dim ctlTest As Control
    For Each ctlTest In F.Controls
        If (ctlTest.Name = ctl.Name) And (Not (ctlTest Is ctl)) Then
            'if the object is the same name but is not the same object we can assume it is a control array
            FindIndex = ctl.Index
            Exit Function
        End If
    Next
    'if we get here then no controls on the form have the same name so can't be a control array
    FindIndex = -1
End Function

Public Sub GetHeaders1(Frm As Form)
    On Error Resume Next
    Dim Txt    As TextBox
    Dim lbl    As Label
    Dim xTab   As C1Tab
    Dim Temp   As C1Tab
    Dim ctl    As Control
    Dim ConCtl As Control
    Dim i      As Integer
    Dim ss     As String
    ss = ""
 
    For Each ctl In Frm.Controls
 
        If TypeOf ctl Is Label Then
            If ctl.Visible = True And Trim(ctl.Caption) <> "" Then
                If FindIndex(Frm, ctl) = -1 Then
                    'ss = ss & Ctl.name & "* " & Ctl.Caption & "* " & "XX" & "* " & vbCrLf
                    ss = ss & ctl.Name & "#" & -1 & " * " & ctl.Caption & " * " & "XX" & vbCrLf
                Else
                    ss = ss & ctl.Name & "#" & ctl.Index & " * " & ctl.Caption & " * " & "XX" & vbCrLf
                End If
                                
                '      ss = ss & ctl.name & "* " & ctl.Caption & "* " & "XX" & "* " & vbCrLf
                CreatLog_File_for_GetHeaders ss, Frm.Name, ss
            End If
        End If

    Next ctl

End Sub

Public Sub GetHeaders(Frm As Form)
    On Error Resume Next
    Dim Txt    As TextBox
    Dim lbl    As Label
    Dim xTab   As C1Tab
    Dim Temp   As C1Tab
    Dim ctl    As Control
    Dim ConCtl As Control
    Dim i      As Integer
    Dim ss     As String
    ss = ""
 
    For Each ctl In Frm.Controls
 
        If TypeOf ctl Is Label Then
            If ctl.Visible = True And Trim(ctl.Caption) <> "" Then
                '                  ss = ss & " " & ctl.name & "#" & ctl.Index & "*" & ctl.Caption & "*" & vbCrLf
                ss = ss & ctl.Name & "* " & ctl.Caption & "* " & "XX" & "* " & vbCrLf
                CreatLog_File_for_GetHeaders ss, Frm.Name, ss
            End If
        End If

    Next ctl

End Sub
Public Sub ShowFormtitles(Frm As Form)
    Dim My_SQL   As String
    Dim Msg      As String
    Dim FileName As String
    FileName = App.path & "\titles\" & Frm.Name & ".txt"
 
    If Dir(FileName, vbNormal) = "" Then
        Msg = " File No found  ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
        Exit Sub
    End If
        
    Open FileName For Input As #1
 
    Dim a               As String
    Dim VarSet()        As String
    Dim Label()         As String
    Dim controlname     As String
    Dim controlindex    As String
    Dim ControlCaptionA As String
    Dim ControlCaptionE As String
    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
                ControlCaptionA = Trim(VarSet(1))
                ControlCaptionE = Trim(VarSet(2))
                Label = Split(VarSet(0), "#", , vbTextCompare)
                If (Label(0)) <> Empty Or Trim(Label(0)) <> "" Then
                    controlname = Trim(Label(0))
                
                End If
 
                Dim Txt    As TextBox
                Dim lbl    As Label
                Dim xTab   As C1Tab
                Dim Temp   As C1Tab
                Dim ctl    As Control
                Dim ConCtl As Control
                Dim i      As Integer
                Dim ss     As String
                ss = ""
 
                'On Error Resume Next
                For Each ctl In Frm.Controls
 
                    If TypeOf ctl Is Label Then
                    
                        If Trim(ctl.Name) = controlname Then
                                                      
                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                      
                                ctl.Caption = ControlCaptionA
                            Else
                                                        
                                ctl.Caption = ControlCaptionE
                            End If
                        End If
                              
                    End If

                Next ctl
               
            End If
        End If
    
    Loop

    Close #1

End Sub
Public Sub ShowFormtitles1(Frm As Form)
    Dim My_SQL   As String
    Dim Msg      As String
    Dim FileName As String
    FileName = App.path & "\titles\" & Frm.Name & ".txt"
 
    If Dir(FileName, vbNormal) = "" Then
        Msg = " File No found  ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
        Exit Sub
    End If
        
    Open FileName For Input As #1
 
    Dim a               As String
    Dim VarSet()        As String
    Dim Label()         As String
    Dim controlname     As String
    Dim controlindex    As String
    Dim ControlCaptionA As String
    Dim ControlCaptionE As String
    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
                ControlCaptionA = Trim(VarSet(1))
                ControlCaptionE = Trim(VarSet(2))
                Label = Split(VarSet(0), "#", , vbTextCompare)
                If (Label(0)) <> Empty Or Trim(Label(0)) <> "" Then
                    controlname = Trim(Label(0))
                    If (Label(1)) <> Empty Then
                        controlindex = val(Label(1))
                    Else
                        controlindex = -1
                    End If
            
                End If
 
                Dim Txt    As TextBox
                Dim lbl    As Label
                Dim xTab   As C1Tab
                Dim Temp   As C1Tab
                Dim ctl    As Control
                Dim ConCtl As Control
                Dim i      As Integer
                Dim ss     As String
                ss = ""
 
                'On Error Resume Next
                For Each ctl In Frm.Controls
 
                    If TypeOf ctl Is Label Then
                    
                        If Trim(ctl.Name) = controlname Then
                                                      
                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                  
                                If controlindex = -1 Then
                                    ctl.Caption = ControlCaptionA
                                Else
                                    '   Ctl.Index = controlindex
                                    ctl.Caption = ControlCaptionA
                                    '           Ctl(controlindex).Caption = ControlCaptionA
                                    '    Ctl( Ctl.name & i).Caption =
                                                                                     
                                End If
                                                                       
                            Else
                                If controlindex = -1 Then
                                    ctl.Caption = ControlCaptionE
                                Else
                                    ctl.Caption = ControlCaptionE
                                End If
                            End If
                        End If
                              
                    End If

                Next ctl
               
            End If
        End If
    
    Loop

    Close #1

End Sub
 
Public Sub SetInterface(Frm As Form)
    On Error Resume Next
    Dim Txt    As TextBox
    Dim lbl    As Label
    Dim xTab   As C1Tab
    Dim Temp   As C1Tab
    Dim ctl    As Control
    Dim ConCtl As Control
    Dim i      As Integer
    'Dim dd As VSFlex8Ctl.VSFlexGrid
    Frm.RightToLeft = False

    For Each ConCtl In Frm.Controls

        If TypeOf ConCtl Is C1Tab Then
            If ConCtl.Position = tpRightHorz Then
                ConCtl.Position = tpLeftHorz
            End If

            ConCtl.CaptionPos = cpLeftCenter

            If TypeOf ConCtl.Container Is Form Then
                ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
            Else
                ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
            End If

        ElseIf TypeOf ConCtl Is frame Then
            ConCtl.RightToLeft = False
            ConCtl.top = ConCtl.top

            If TypeOf ConCtl.Container Is Form Then
                ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
            Else
                ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
            End If

        ElseIf TypeOf ConCtl Is C1Elastic Then

            If Not TypeOf ConCtl.Container Is C1Tab Then
                If TypeOf ConCtl.Container Is Form Then
                    ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
                Else
                    ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
                End If
            End If

            If ConCtl.CaptionPos = cpRightCenter Then
                ConCtl.CaptionPos = cpLeftCenter
            ElseIf ConCtl.CaptionPos = cpRightTop Then
                ConCtl.CaptionPos = cpLeftTop
            End If

            If ConCtl.PicturePos = ppLeftCenter Then
                ConCtl.PicturePos = ppRightCenter
            ElseIf ConCtl.PicturePos = ppLeftTop Then
                ConCtl.PicturePos = ppRightTop
            End If

        ElseIf TypeOf ConCtl Is PictureBox Then
            ConCtl.RightToLeft = False

            If TypeOf ConCtl.Container Is Form Then
                ConCtl.left = ConCtl.Container.ScaleWidth - (ConCtl.left + ConCtl.Width)
            Else
                ConCtl.left = ConCtl.Container.Width - (ConCtl.left + ConCtl.Width)
            End If
        End If

    Next ConCtl

    For Each ctl In Frm.Controls

        If Not (TypeOf ctl Is Timer Or TypeOf ctl Is CommonDialog Or TypeOf ctl Is line Or TypeOf ctl Is ImageList Or TypeOf ctl Is frame Or TypeOf ctl Is C1Elastic Or TypeOf ctl Is C1Tab Or TypeOf ctl Is Menu Or TypeOf ctl Is PictureBox Or TypeOf ctl Is XPPopUp30) Then

            If TypeOf ctl Is TextBox Then
                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If
                If ctl.Alignment <> vbCenter Then
                    ctl.RightToLeft = False
                End If
                
            ElseIf TypeOf ctl Is Label Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is ISAniLabel Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is DTPicker Then

                If ctl.Format = dtpCustom Then
                    '    Ctl.CustomFormat = "d/M/yyyy"
                End If

            ElseIf TypeOf ctl Is DataCombo Then
                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is CheckBox Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is OptionButton Then

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is CommandButton Then
                ctl.RightToLeft = False
            ElseIf (TypeOf ctl Is ISButton) Or (TypeOf ctl Is ISButtonLW) Then
                ctl.RightToLeft = False

                If ctl.Alignment = vbRightJustify Then
                    ctl.Alignment = vbLeftJustify
                End If

                If ctl.ButtonPositionImage = impRightOfText Then
                    ctl.ButtonPositionImage = impLeftOfText
                End If

            ElseIf TypeOf ctl Is ComboBox Then
                ctl.RightToLeft = False
            ElseIf TypeOf ctl Is Image Then
            ElseIf (TypeOf ctl Is VSFlex8Ctl.VSFlexGrid) Or (TypeOf ctl Is VSFlex8UCtl.VSFlexGrid) Then
                ctl.RightToLeft = False

                For i = 0 To ctl.Cols - 1

                    If ctl.ColAlignment(i) = flexAlignRightCenter Then
                        ctl.ColAlignment(i) = flexAlignLeftCenter
                    End If

                    If ctl.FixedAlignment(i) = flexAlignRightCenter Then
                        ctl.FixedAlignment(i) = flexAlignLeftCenter
                    End If

                Next i

                'ElseIf TypeOf Ctl Is AniGIF Then
            
            End If

            'If Ctl.Name = "CBtnColor" Then Stop
            If TypeOf ctl.Container Is Form Then
                ctl.left = ctl.Container.ScaleWidth - (ctl.left + ctl.Width)
            Else
                ctl.left = ctl.Container.Width - (ctl.left + ctl.Width)
            End If
        End If

    Next ctl

End Sub

Public Sub CreatLogFile()
    Dim StrLogFileName As String
    Dim IntFreeFile    As Integer
    Dim ss             As String

    StrLogFileName = App.path & "\LoadLog.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "Log File For  Dynamic Byte System"
    ss = ss & vbCrLf & "Stars Tech. "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Public Sub WriteInLogFile(ss As String)
    Dim StrLogFileName As String
    Dim IntFreeFile    As Integer
    Dim StrTemp        As String

    StrLogFileName = App.path & "\LoadLog.txt"

    If Dir(StrLogFileName) = "" Then
        CreatLogFile
    End If

    IntFreeFile = FreeFile
    StrTemp = String(10, "-")
    StrTemp = StrTemp & vbCrLf & Now
    StrTemp = StrTemp & vbCrLf & ss
    Open StrLogFileName For Append As #IntFreeFile
    Write #IntFreeFile, StrTemp
    Close #IntFreeFile

End Sub

Public Sub OpenFile(StrFilePath As String)
    On Error Resume Next

    If Dir(StrFilePath) <> "" Then
        ShellExecute mdifrmmain.hWnd, vbNullString, StrFilePath, vbNullString, "", SW_SHOWNORMAL
    End If

End Sub

Public Function GetMyTitleBarHight(FrmHwnd As Long) As Single
    Dim TitleInfo As TITLEBARINFO
    Dim SngTemp   As Single

    'Initialize structure
    TitleInfo.cbSize = Len(TitleInfo)
    'Retrieve information about the tilte bar of this window
    GetTitleBarInfo FrmHwnd, TitleInfo
    'Show some of that information
    SngTemp = TitleInfo.rcTitleBar.bottom '- TitleInfo.rcTitleBar.Top
    'SngTemp = SngTemp * Screen.TwipsPerPixelY
    GetMyTitleBarHight = SngTemp

    'Debug.Print "Title bar rectangle:"
    'With TitleInfo.rcTitleBar
    '        Debug.Print "   (" & CStr(.Left) & "," & CStr(.Top) & ")-(" & CStr(.Right) & "," & CStr(.Bottom) & ")"
    'End With

End Function

Public Sub PutFormOnTop(LngHwnd As Long, _
   Optional BolOnTop As Boolean = True)

    If BolOnTop = True Then
        SetWindowPos LngHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos LngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If

End Sub

Public Function GetAppTitle() As String
    Dim StrTemp       As String
    Dim StrRegCaption As String

    If SystemOptions.SysRegisterState = DemoRun Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrRegCaption = "‰”Œ…  Ã—Ì»Ì…"
        Else
            StrRegCaption = "Demo Version"
        End If

    ElseIf SystemOptions.SysRegisterState = DemoStop Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrRegCaption = "‰”Œ…  Ã—Ì»Ì… „‰ ÂÌ…"
        Else
            StrRegCaption = "Demo Expired"
        End If

    ElseIf SystemOptions.SysRegisterState = DevelopVersion Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrRegCaption = "‰”Œ… „”Ã·…  "
        Else
            StrRegCaption = "Develop Version"
        End If

    ElseIf SystemOptions.SysRegisterState = Registered Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrRegCaption = "‰”Œ… „”Ã·…"
        Else
            StrRegCaption = "Registered Version"
        End If

    ElseIf SystemOptions.SysRegisterState = UnErrorOccured Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrRegCaption = "‰”Œ… €Ì— √’·Ì…"
        Else
            StrRegCaption = "Cracked Version"
        End If
    End If

    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        If SystemOptions.UserInterface = ArabicInterface Then
            '  StrTemp = "  »«Ì  " & StrRegCaption & App.Major & " :  " & App.Minor & " :  " & App.Revision
            StrTemp = "  »«Ì  " & StrRegCaption & getLastDataBaseUpdateDate
        Else
            '        StrTemp = "Byte For Complete Accounting   " & StrRegCaption & App.Major & " : " & App.Minor & " :  " & App.Revision
            StrTemp = "Byte For Complete Accounting   " & getLastDataBaseUpdateDate
        End If

    ElseIf SystemOptions.SysAppAccoutingType = SimpleAccoutning Then

        If SystemOptions.UserInterface = ArabicInterface Then
            'StrTemp = "‰Ÿ«„ œÌ‰«„Ìﬂ »«Ì   «·„ ﬂ«„· " & StrRegCaption & App.Major & ":" & App.Minor & ":" & App.Revision
            StrTemp = "  »«Ì  " & StrRegCaption & getLastDataBaseUpdateDate
        Else
            'StrTemp = "Byte For Accounting " & StrRegCaption & App.Major & ":" & App.Minor & ":" & App.Revision
            StrTemp = "Byte For Complete Accounting   " & getLastDataBaseUpdateDate
        End If
    End If

    ''GetAppTitle = getLastDataBaseUpdateDate & "   " & CurrentActivityName & "   " & CurrentBranchName
    'GetAppTitle = GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·—∆Ì”Ì…") & "- " & CurrentActivityName & "   " & CurrentBranchName
    If SystemOptions.UserInterface = ArabicInterface Then
        GetAppTitle = "«·ﬁ«⁄œ… «·Õ«·Ì…: " & GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·—∆Ì”Ì…")   '& "- " & CurrentActivityName & "   " & CurrentBranchName
    Else
        GetAppTitle = "Current DataBase Name: " & GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·—∆Ì”Ì…")   '& "- " & CurrentActivityName & "   " & CurrentBranchName
    End If
End Function

Public Function GetTaskBarHeight() As Long
    Dim DesktophWnd   As Long
    Dim r             As RECT
    Dim tWnd          As Long
    Dim TaskBarTop    As Long
    Dim TaskBarRight  As Long
    Dim TaskBarHeight As Long
    Dim TaskBarBottom As Long

    On Error GoTo ErrTrap
    'get the desk top window handle
    DesktophWnd = GetDesktopWindow
    'get the Desktop Window dimention
    GetWindowRect DesktophWnd, r
    'Get the taskbar's window handle
    'Note the "Shell_TrayWnd" is the Class name for the taskbar
    tWnd = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect tWnd, r

    TaskBarTop = r.top
    TaskBarRight = r.right
    TaskBarBottom = r.bottom

    TaskBarTop = TaskBarTop * Screen.TwipsPerPixelX
    TaskBarBottom = TaskBarBottom * Screen.TwipsPerPixelX
    TaskBarHeight = TaskBarBottom - TaskBarTop
    GetTaskBarHeight = TaskBarHeight
    Exit Function
ErrTrap:
    GetTaskBarHeight = 0
End Function

Public Function CorrectCurrency(SngNumber As Single) As Currency
    Dim IntIntgerFactor As Integer
    Dim IntFractional   As Single
    Dim SngMod          As Single

    IntFractional = SngNumber - Int(SngNumber)
    IntFractional = IntFractional * 100
    IntFractional = Format(IntFractional, SystemOptions.SysDefCurrencyForamt)
    SngMod = IntFractional Mod 5

    If SngMod <> 0 Then
        If SngMod = 1 Or SngMod = 2 Or SngMod = 3 Then
            IntFractional = IntFractional - SngMod
            CorrectCurrency = Int(SngNumber) + (IntFractional / 100)
        ElseIf SngMod = 4 Or SngMod = 5 Then
            IntFractional = IntFractional + (IIf(SngMod = 4, 1, 0))
            CorrectCurrency = Int(SngNumber) + (IntFractional / 100)
        End If

    Else
        CorrectCurrency = SngNumber
    End If

End Function

Public Function ShowDynamicHelp(Lngid As Long) As Boolean
    Dim StrFile As String
    On Error GoTo ErrTrap
    StrFile = App.path & "\DynamicHelp\" & Lngid & ".htm"

    If Dir(StrFile, vbNormal) <> "" Then
        If Not FrmDynamicHelpPane Is Nothing Then

            'FrmDynamicHelpPane.WbHelp.Navigate StrFile
            If FrmDynamicHelpPane.WbHelp.Busy = True Then
                FrmDynamicHelpPane.WbHelp.Stop
            End If

            FrmDynamicHelpPane.WbHelp.Navigate StrFile

            DoEvents
            FrmDynamicHelpPane.WbHelp.Refresh
        End If
    End If

    Exit Function
ErrTrap:
End Function

Public Function LoadMainSystemOptions() As Boolean
    Dim rs       As ADODB.Recordset
    Dim Msg      As String
    Dim IntTemp  As Integer
    Dim IntTemp1 As Integer

    Dim StrTemp  As String
    Dim StrSQL   As String

    'On Error GoTo hErr
    On Error Resume Next
    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    SystemOptions.SysAllowStockNegative = IIf(rs("AllowStockNegative").value = 0 Or IsNull(rs("AllowStockNegative").value), False, True)
    
    SystemOptions.NotAllowStockNegativeInternal = IIf(rs("NotAllowStockNegativeInternal").value = 0 Or IsNull(rs("NotAllowStockNegativeInternal").value), False, True)
    SystemOptions.MustEnterNewNo = IIf(rs("MustEnterNewNo").value = 0 Or IsNull(rs("MustEnterNewNo").value), False, True)

    SystemOptions.SysAllowBoxNegative = IIf(rs("AllowBoxNegative").value = 0 Or IsNull(rs("AllowBoxNegative").value), False, True)
    SystemOptions.SysMantainceAllow = IIf(rs("MantainceAllow").value = 0 Or IsNull(rs("MantainceAllow").value), False, True)
    
    SystemOptions.SysMainStockCostMethod = rs("MainStockCostType").value
    SystemOptions.itemSeprator = rs("itemSeprator").value
    SystemOptions.DefaultQtyTrans = IIf(IsNull(rs("DefaultQtyTrans").value), 1, rs("DefaultQtyTrans").value)
    
    ''''
    SystemOptions.ChasingStatus = IIf(rs("ChasingStatus").value = 0 Or IsNull(rs("ChasingStatus").value), 1, rs("ChasingStatus").value)
    ''''
    
    SystemOptions.Items_or_operation = IIf(IsNull(rs("Items_or_operation").value), -1, val(rs("Items_or_operation").value))
    
    SystemOptions.ProjectDiscountPolicy = IIf(IsNull(rs("ProjectDiscountPolicy").value), 0, (rs("ProjectDiscountPolicy").value))
    
    'ProjectDiscountPolicy
  
    SystemOptions.gldetails_or_gl_general = IIf(IsNull(rs("gl_detaila_or_total").value), -1, val(rs("gl_detaila_or_total").value))
    SystemOptions.ProcessPeriodType = IIf(IsNull(rs("ProcessPeriodType").value), -1, val(rs("ProcessPeriodType").value))

    SystemOptions.itemcodePart1 = IIf(IsNull(rs("itemcodePart1").value), -1, val(rs("itemcodePart1").value))
    SystemOptions.itemcodePart2 = IIf(IsNull(rs("itemcodePart2").value), -1, val(rs("itemcodePart2").value))
    SystemOptions.itemcodePart3 = IIf(IsNull(rs("itemcodePart3").value), -1, val(rs("itemcodePart3").value))
    SystemOptions.itemcodePart1NoOFDigit = IIf(IsNull(rs("itemcodePart1NoOFDigit").value), 0, val(rs("itemcodePart1NoOFDigit").value))
    SystemOptions.itemcodePart2NoOFDigit = IIf(IsNull(rs("itemcodePart2NoOFDigit").value), 0, val(rs("itemcodePart2NoOFDigit").value))
    SystemOptions.itemcodePart3NoOFDigit = IIf(IsNull(rs("itemcodePart3NoOFDigit").value), 0, val(rs("itemcodePart3NoOFDigit").value))
    SystemOptions.itemcodeSeperator1 = IIf(IsNull(rs("itemcodeSeperator1").value), "", (rs("itemcodeSeperator1").value))
    SystemOptions.itemcodeSeperator2 = IIf(IsNull(rs("itemcodeSeperator2").value), "", (rs("itemcodeSeperator2").value))
    SystemOptions.itemsWorkWithSize = IIf(rs("itemsWorkWithSize").value = 0 Or IsNull(rs("itemsWorkWithSize").value), False, True)
  
    SystemOptions.itemsWorkWithColor = IIf(rs("itemsWorkWithColor").value = 0 Or IsNull(rs("itemsWorkWithColor").value), False, True)
    SystemOptions.itemsWorkWithDates = IIf(rs("itemsWorkWithDates").value = 0 Or IsNull(rs("itemsWorkWithDates").value), False, True)
    SystemOptions.itemsWorkWithClass = IIf(rs("itemsWorkWithClass").value = 0 Or IsNull(rs("itemsWorkWithClass").value), False, True)

    SystemOptions.ItemcodeGroupOnly = IIf(rs("ItemcodeGroupOnly").value = 0 Or IsNull(rs("ItemcodeGroupOnly").value), False, True)
    SystemOptions.ItemcodeGroupandParentGroup = IIf(rs("ItemcodeGroupandParentGroup").value = 0 Or IsNull(rs("ItemcodeGroupandParentGroup").value), False, True)
 
    SystemOptions.SaleDiscount1 = IIf(IsNull(rs("SaleDiscount1").value), 0, val(rs("SaleDiscount1").value))
    SystemOptions.SaleDiscount2 = IIf(IsNull(rs("SaleDiscount2").value), 0, val(rs("SaleDiscount2").value))
    SystemOptions.SaleDiscount3 = IIf(IsNull(rs("SaleDiscount3").value), 0, val(rs("SaleDiscount3").value))
    SystemOptions.SaleDiscount4 = IIf(IsNull(rs("SaleDiscount4").value), 0, val(rs("SaleDiscount4").value))
    
    SystemOptions.autoIssueVoucher = IIf(rs("autoIssueVoucher").value = 0 Or IsNull(rs("autoIssueVoucher").value), False, True)
 
    SystemOptions.MonthIs30days = IIf(rs("MonthIs30days").value = 0 Or IsNull(rs("MonthIs30days").value), False, True)
    SystemOptions.RawMaterMix = IIf(rs("RawMaterMix").value = 0 Or IsNull(rs("RawMaterMix").value), False, True)
    SystemOptions.bankComm = IIf(rs("bankComm").value = 0 Or IsNull(rs("bankComm").value), False, True)
    SystemOptions.ChequeBox = IIf(rs("ChequeBox").value = 0 Or IsNull(rs("ChequeBox").value), False, True)
    SystemOptions.IsCheque = IIf(rs("IsCheque").value = 0 Or IsNull(rs("IsCheque").value), False, True)
    
    SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").value = 0 Or IsNull(rs("CustomerhavethreeAccounts").value), False, True)
    SystemOptions.IsCreateOpenBalnceMan = IIf(rs("IsCreateOpenBalnceMan").value = 0 Or IsNull(rs("IsCreateOpenBalnceMan").value), False, True)
    
    SystemOptions.CustomerhavethreeAccounts1 = IIf(rs("CustomerhavethreeAccounts1").value = 0 Or IsNull(rs("CustomerhavethreeAccounts1").value), False, True)
  
    SystemOptions.AllowRepeateCar = IIf(rs("AllowRepeateCar").value = 0 Or IsNull(rs("AllowRepeateCar").value), False, True)
    SystemOptions.CostByProduction = IIf(rs("CostByProduction").value = 0 Or IsNull(rs("CostByProduction").value), False, True)

    SystemOptions.IsByNewCoding = IIf(rs("IsByNewCoding").value = 0 Or IsNull(rs("IsByNewCoding").value), False, True)
    SystemOptions.IsAutoNameItems = IIf(rs("IsAutoNameItems").value = 0 Or IsNull(rs("IsAutoNameItems").value), False, True)
    SystemOptions.mDomainData = rs!DomainData & ""
    APIURL = SystemOptions.mDomainData
    'modmod
    SystemOptions.PaymentMethLaterCompItem = IIf(rs("PaymentMethLaterCompItem").value = 0 Or IsNull(rs("PaymentMethLaterCompItem").value), False, True)
    SystemOptions.ShowBalanceCustInv = IIf(rs("ShowBalanceCustInv").value = 0 Or IsNull(rs("ShowBalanceCustInv").value), False, True)
  
    SystemOptions.MaintOrderCantRepeatSales = IIf(rs("MaintOrderCantRepeatSales").value = 0 Or IsNull(rs("MaintOrderCantRepeatSales").value), False, True)
    SystemOptions.MaintOrderCantRepeatBillBuy = IIf(rs("MaintOrderCantRepeatBillBuy").value = 0 Or IsNull(rs("MaintOrderCantRepeatBillBuy").value), False, True)

    SystemOptions.TripRevenueAuto = IIf(rs("TripRevenueAuto").value = 0 Or IsNull(rs("TripRevenueAuto").value), False, True)

    SystemOptions.cdoSMTPUseSSL = IIf(rs("cdoSMTPUseSSL").value = 0 Or IsNull(rs("cdoSMTPUseSSL").value), False, True)

    SystemOptions.cdoSMTPServerPort = IIf(rs("cdoSMTPServerPort").value = 0 Or IsNull(rs("cdoSMTPServerPort").value), 587, rs("cdoSMTPServerPort").value)

    SystemOptions.cdoSMTPServer = IIf(rs("cdoSMTPServer").value = "" Or IsNull(rs("cdoSMTPServer").value), "cdoSMTPServer", rs("cdoSMTPServer").value)

    SystemOptions.cdoSendUserName = IIf(rs("cdoSendUserName").value = "" Or IsNull(rs("cdoSendUserName").value), "cdoSendUserName", rs("cdoSendUserName").value)
    SystemOptions.cdoSendPassword = IIf(rs("cdoSendPassword").value = "" Or IsNull(rs("cdoSendPassword").value), "cdoSendPassword", rs("cdoSendPassword").value)

    SystemOptions.TxtFromName = IIf(rs("txtFromName").value = "" Or IsNull(rs("txtFromName").value), " ", rs("txtFromName").value)
    SystemOptions.txtFromEmail = IIf(rs("txtFromEmail").value = "" Or IsNull(rs("txtFromEmail").value), " ", rs("txtFromEmail").value)

    SystemOptions.NoOFDigitUserTrans = IIf(val(rs!IsSerialByUserTrans & "") = 0, 2, val(rs!IsSerialByUserTrans & ""))
    SystemOptions.NoOFDigitUserVouc = IIf(val(rs!NoOFDigitUserVouc & "") = 0, 2, val(rs!NoOFDigitUserVouc & ""))

    SystemOptions.IsSerialByUserTrans = IIf(rs("IsSerialByUserTrans").value = 0 Or IsNull(rs("IsSerialByUserTrans").value), False, True)
    SystemOptions.IsSerialByUserVouch = IIf(rs("IsSerialByUserVouch").value = 0 Or IsNull(rs("IsSerialByUserVouch").value), False, True)

    SystemOptions.IsSomeItemWeight = IIf(rs("IsSomeItemWeight").value = 0 Or IsNull(rs("IsSomeItemWeight").value), False, True)
    SystemOptions.IsMergeVat = IIf(rs("IsMergeVat").value = 0 Or IsNull(rs("IsMergeVat").value), False, True)
  
    SystemOptions.FromNo = val(rs("FromNo") & "")
    'val(IIf(IsNull(rs("FromNo").value), 0, val(rs("FromNo").value)) & "")
    SystemOptions.OrNo = val(rs("OrNo").value & "")  'IIf(IsNull(rs("OrNo").value), 0, val(rs("OrNo").value))
    SystemOptions.CodeFrom = val(rs("CodeFrom").value & "")   'IIf(IsNull(rs("CodeFrom").value), 0, val(rs("CodeFrom").value))
    SystemOptions.CodeTo = val(rs("CodeTo").value & "")    ' IIf(IsNull(rs("CodeTo").value), 0, val(rs("CodeTo").value))
    SystemOptions.WeightFrom = val(rs("WeightFrom").value & "")    ' IIf(IsNull(rs("WeightFrom").value), 0, val(rs("WeightFrom").value))
    SystemOptions.WeightTo = val(rs("WeightTo").value & "")    ' IIf(IsNull(rs("WeightTo").value), 0, val(rs("WeightTo").value))

    SystemOptions.IsGeometricProportions = IIf(rs("IsGeometricProportions").value = 0 Or IsNull(rs("IsGeometricProportions").value), False, True)
  
    SystemOptions.CanPartialpayment = IIf(rs("CanPartialpayment").value = 0 Or IsNull(rs("CanPartialpayment").value), False, True)
  
    SystemOptions.EndRentifPayed = IIf(rs("EndRentifPayed").value = 0 Or IsNull(rs("EndRentifPayed").value), False, True)
    SystemOptions.cantCahngeAkarinExpenses = IIf(rs("cantCahngeAkarinExpenses").value = 0 Or IsNull(rs("cantCahngeAkarinExpenses").value), False, True)
  
    SystemOptions.EmployeeSalaryBYBranch = IIf(rs("EmployeeSalaryBYBranch").value = 0 Or IsNull(rs("EmployeeSalaryBYBranch").value), False, True)
    SystemOptions.returnnotcreatvoucher = IIf(rs("returnnotcreatvoucher").value = 0 Or IsNull(rs("returnnotcreatvoucher").value), False, True)
    SystemOptions.OnlyOneCashingVchr = IIf(rs("OnlyOneCashingVchr").value = 0 Or IsNull(rs("OnlyOneCashingVchr").value), False, True)
  
    SystemOptions.CheckDateFormatCorrect = IIf(rs("CheckDateFormatCorrect").value = 0 Or IsNull(rs("CheckDateFormatCorrect").value), False, True)
    SystemOptions.CheckMobileFormatCorrect = IIf(rs("CheckMobileFormatCorrect").value = 0 Or IsNull(rs("CheckMobileFormatCorrect").value), False, True)
  
    SystemOptions.CantRepetttransferNoforCashing = IIf(rs("CantRepetttransferNoforCashing").value = 0 Or IsNull(rs("CantRepetttransferNoforCashing").value), False, True)
   
    SystemOptions.WaiverSetByContract = IIf(rs("WaiverSetByContract").value = 0 Or IsNull(rs("WaiverSetByContract").value), False, True)
    SystemOptions.WaiverSetByContract = IIf(rs("WaiverSetByContract").value = 0 Or IsNull(rs("WaiverSetByContract").value), False, True)
  
    SystemOptions.CarsRevenuePerOwner = IIf(rs("CarsRevenuePerOwner").value = 0 Or IsNull(rs("CarsRevenuePerOwner").value), False, True)
  
    SystemOptions.IsCustSalesManCashRelated = IIf(rs("IsCustSalesManCashRelated").value = 0 Or IsNull(rs("IsCustSalesManCashRelated").value), False, True)
    SystemOptions.showEmployeeAccountIntrip = IIf(rs("showEmployeeAccountIntrip").value = 0 Or IsNull(rs("showEmployeeAccountIntrip").value), False, True)
    SystemOptions.DUEDOCUMENTbyinstallDate = IIf(rs("DUEDOCUMENTbyinstallDate").value = 0 Or IsNull(rs("DUEDOCUMENTbyinstallDate").value), False, True)
  
    SystemOptions.CanSkipPurchOrder = IIf(rs("CanSkipPurchOrder").value = 0 Or IsNull(rs("CanSkipPurchOrder").value), False, True)
  
    SystemOptions.CompilingBasedTable = IIf(rs("CompilingBasedTable").value = 0 Or IsNull(rs("CompilingBasedTable").value), False, True)
    SystemOptions.DontSaveInvoiceWithoutDocType = IIf(rs("DontSaveInvoiceWithoutDocType").value = 0 Or IsNull(rs("DontSaveInvoiceWithoutDocType").value), False, True)
  
    SystemOptions.DontDuplicateManulaNoInPurchase = IIf(rs("DontDuplicateManulaNoInPurchase").value = 0 Or IsNull(rs("DontDuplicateManulaNoInPurchase").value), False, True)
    SystemOptions.SpecialVersion = IIf(rs("SpecialVersion").value = 0 Or IsNull(rs("SpecialVersion").value), False, True)
  
    SystemOptions.InvoiceTransferJLTotal = IIf(rs("InvoiceTransferJLTotal").value = 0 Or IsNull(rs("InvoiceTransferJLTotal").value), False, True)
  
    SystemOptions.EmpProduction = IIf(rs("EmpProduction").value = 0 Or IsNull(rs("EmpProduction").value), False, True)
    SystemOptions.ItemProduction = IIf(rs("ItemProduction").value = 0 Or IsNull(rs("ItemProduction").value), False, True)
    SystemOptions.ExpProduction = IIf(rs("ExpProduction").value = 0 Or IsNull(rs("ExpProduction").value), False, True)
    
    SystemOptions.VATNoAccordActivity = IIf(rs("VATNoAccordActivity").value = 0 Or IsNull(rs("VATNoAccordActivity").value), False, True)
    SystemOptions.NotCrtResvVouchProjects = IIf(rs("NotCrtResvVouchProjects").value = 0 Or IsNull(rs("NotCrtResvVouchProjects").value), False, True)
    SystemOptions.IsShowLensesDetails = IIf(rs("IsShowLensesDetails").value = 0 Or IsNull(rs("IsShowLensesDetails").value), False, True)

    SystemOptions.LinkUsersWithPayment = IIf(rs("LinkUsersWithPayment").value = 0 Or IsNull(rs("LinkUsersWithPayment").value), False, True)

    SystemOptions.logowidth = IIf(rs("logowidth").value = 0 Or IsNull(rs("logowidth").value), 4000, rs("logowidth").value)
    SystemOptions.logoHeight = IIf(rs("logoHeight").value = 0 Or IsNull(rs("logoHeight").value), 1500, rs("logoHeight").value)
 
    SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").value = 0 Or IsNull(rs("CustomerhavethreeAccounts").value), False, True)
    SystemOptions.IsCreateOpenBalnceMan = IIf(rs("IsCreateOpenBalnceMan").value = 0 Or IsNull(rs("IsCreateOpenBalnceMan").value), False, True)
 
    SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").value = 0 Or IsNull(rs("CustomerhavethreeAccounts").value), False, True)
 
    SystemOptions.CreateDriverBox = IIf(rs("CreateDriverBox").value = 0 Or IsNull(rs("CreateDriverBox").value), False, True)
    SystemOptions.CreateDriverEra = IIf(rs("CreateDriverEra").value = 0 Or IsNull(rs("CreateDriverEra").value), False, True)
 
    SystemOptions.TypicalProduction = IIf(rs("TypicalProduction").value = 0 Or IsNull(rs("TypicalProduction").value = 0), False, True)
 
    SystemOptions.ExpensesCoding = IIf(rs("ExpensesCoding").value = 0 Or IsNull(rs("ExpensesCoding").value), False, True)
    SystemOptions.ExpensesCoding2 = IIf(rs("ExpensesCoding2").value = 0 Or IsNull(rs("ExpensesCoding2").value), False, True)
    SystemOptions.SMSUserName = IIf(IsNull(rs("SMSUserName").value), "", rs("SMSUserName").value)
    SystemOptions.SMSPassWord = IIf(IsNull(rs("SMSPassWord").value), "", rs("SMSPassWord").value)
    SystemOptions.SenderName = IIf(IsNull(rs("SenderName").value), "", rs("SenderName").value)
    SystemOptions.OPTWEB = IIf(IsNull(rs("optweb").value), 0, rs("optweb").value)
   
    SystemOptions.InstallmntsvchrCoding = IIf(rs("InstallmntsvchrCoding").value = 0 Or IsNull(rs("InstallmntsvchrCoding").value), False, True)
 
    SystemOptions.AllowIndirectCost = IIf(rs("AllowIndirectCost").value = 0 Or IsNull(rs("AllowIndirectCost").value), False, True)
 
    SystemOptions.banks_Accounts3 = IIf(rs("banks_Accounts").value = 0 Or IsNull(rs("banks_Accounts").value), False, True)
    SystemOptions.AssetAccount = IIf(rs("AssetAccount").value = 0 Or IsNull(rs("AssetAccount").value), False, True)
    SystemOptions.AssetAccount1 = IIf(rs("AssetAccount1").value = 0 Or IsNull(rs("AssetAccount1").value), False, True)
    '**********************************************************************
    SystemOptions.StoreAccountHaveSettelment = IIf(rs("StoreAccountHaveSettelment").value = 0 Or IsNull(rs("StoreAccountHaveSettelment").value), False, True)
    SystemOptions.eachStoreHaveLossAccount = IIf(IsNull(rs("eachStoreHaveLossAccount").value), True, IIf((rs("eachStoreHaveLossAccount").value = 0), False, True))
    SystemOptions.eachStoreHaveGiftAccount = IIf(IsNull(rs("eachStoreHaveGiftAccount").value), True, IIf((rs("eachStoreHaveGiftAccount").value = 0), False, True))
             
    '**********************************************************************
    SystemOptions.autoReseiveVoucher = IIf(rs("autoReseiveVoucher").value = 0 Or IsNull(rs("autoReseiveVoucher").value), False, True)
  
    SystemOptions.ReturnSallingOption = IIf(IsNull(rs("ReturnSallingOption").value), False, (rs("ReturnSallingOption").value))
    SystemOptions.ReturnSallingIntervalCount = IIf(IsNull(rs("ReturnSallingIntervalCount").value), 0, val(rs("ReturnSallingIntervalCount").value))
    SystemOptions.ReturnSallingIntervalCount1 = IIf(IsNull(rs("ReturnSallingIntervalCount1").value), 0, val(rs("ReturnSallingIntervalCount1").value))

    SystemOptions.DateOpt = IIf(IsNull(rs("DateOpt").value), 0, val(rs("DateOpt").value))

    SystemOptions.IndirectCostPercentage = IIf(IsNull(rs("IndirectCostPercentage").value), 0, val(rs("IndirectCostPercentage").value))
    
    SystemOptions.StoreDigit = IIf(IsNull(rs("StoreDigit").value), 1, (rs("StoreDigit").value))
    SystemOptions.BranchDigit = IIf(IsNull(rs("BranchDigit").value), 1, (rs("BranchDigit").value))
    
    SystemOptions.Ked_digit = IIf(IsNull(rs("Ked_digit").value), 0, val(rs("Ked_digit").value))
    SystemOptions.Count_ACCOUNT_digit = IIf(IsNull(rs("Count_ACCOUNT_digit").value), 2, val(rs("Count_ACCOUNT_digit").value))
    SystemOptions.Save_options = IIf(IsNull(rs("Save_options").value), 0, val(rs("Save_options").value))
    SystemOptions.ReservEmp = IIf(IsNull(rs("EmpRes").value), 0, (rs("EmpRes").value))

    SystemOptions.EmpComponentDigts = IIf(IsNull(rs("EmpComponentDigts").value), 2, val(rs("EmpComponentDigts").value))
    SystemOptions.ImagesPath = IIf(IsNull(rs("ImagesPath").value), "Images", (rs("ImagesPath").value))
    SystemOptions.Reportpath = IIf(IsNull(rs("reportPath").value), "Stander", (rs("reportPath").value))
    SystemOptions.BigUserPw = IIf(IsNull(rs("BigUserPw").value), "n20172018", (rs("BigUserPw").value))
    SystemOptions.BigUserPw2 = IIf(IsNull(rs("BigUserPw2").value), "123456", (rs("BigUserPw2").value))
    If SystemOptions.BigUserPw = "" Then SystemOptions.BigUserPw = "n20172018"
    If SystemOptions.BigUserPw2 = "" Then SystemOptions.BigUserPw2 = "123456"
    'BigUserPw
    Report_Folder = SystemOptions.Reportpath

    SystemOptions.CostStarting = IIf(rs("CostStarting").value = 0 Or IsNull(rs("CostStarting").value), False, True)
    SystemOptions.CostStartingGard = IIf(rs("CostStartingGard").value = 0 Or IsNull(rs("CostStartingGard").value), False, True)

    SystemOptions.chkuserCode = IIf(rs("chkuserCode").value = 0 Or IsNull(rs("chkuserCode").value), False, True)
    SystemOptions.Itemsattachedzero = IIf(rs("Itemsattachedzero").value = 0 Or IsNull(rs("Itemsattachedzero").value), False, True)
    SystemOptions.workWithBarcode = IIf(rs("workWithBarcode").value = 0 Or IsNull(rs("workWithBarcode").value), False, True)
    SystemOptions.WorkWithBarCodeParent = IIf(rs("WorkWithBarCodeParent").value = 0 Or IsNull(rs("WorkWithBarCodeParent").value), False, True)

    SystemOptions.amlaketbatrentOnly = IIf(rs("amlaketbatrentOnly").value = 0 Or IsNull(rs("amlaketbatrentOnly").value), False, True)

    SystemOptions.WorkWithLINKEDiActivity = IIf(rs("WorkWithLINKEDiActivity").value = 0 Or IsNull(rs("WorkWithLINKEDiActivity").value), False, True)

    SystemOptions.WorkWithLINKEDiTEMS = IIf(rs("WorkWithLINKEDiTEMS").value = 0 Or IsNull(rs("WorkWithLINKEDiTEMS").value), False, True)
    SystemOptions.WorkWithBranchLogo = IIf(rs("WorkWithBranchLogo").value = 0 Or IsNull(rs("WorkWithBranchLogo").value), False, True)

    SystemOptions.WorkWithFirstInstallOnly = IIf(rs("WorkWithFirstInstallOnly").value = 0 Or IsNull(rs("WorkWithFirstInstallOnly").value), False, True)
    SystemOptions.WorkWithGroupCode = IIf(rs("WorkWithGroupCode").value = 0 Or IsNull(rs("WorkWithGroupCode").value), False, True)
    '31032017egypt

    SystemOptions.AllowSalesMultyPayed = IIf(rs("AllowSalesMultyPayed").value = 0 Or IsNull(rs("AllowSalesMultyPayed").value), False, True)
    SystemOptions.MultyStore = IIf(rs("MultyStore").value = 0 Or IsNull(rs("MultyStore").value), False, True)
    'SystemOptions.MultyStore = False
    SystemOptions.RawMaterMix2 = IIf(rs("RawMaterMix2").value = 0 Or IsNull(rs("RawMaterMix2").value), False, True)
    SystemOptions.DontShowMoreDetailsCompItem = IIf(rs("DontShowMoreDetailsCompItem").value = 0 Or IsNull(rs("DontShowMoreDetailsCompItem").value), False, True)
    SystemOptions.traveDiscountFromCustomerDirect = IIf(rs("traveDiscountFromCustomerDirect").value = 0 Or IsNull(rs("traveDiscountFromCustomerDirect").value), False, True)
 
    SystemOptions.TransferNotInvItemDef = IIf(rs("TransferNotInvItemDef").value = 0 Or IsNull(rs("TransferNotInvItemDef").value), False, True)

    SystemOptions.CustMobNoMandatory = IIf(rs("CustMobNoMandatory").value = 0 Or IsNull(rs("CustMobNoMandatory").value), False, True)
    SystemOptions.CustVatNoMandatory = IIf(rs("CustVatNoMandatory").value = 0 Or IsNull(rs("CustVatNoMandatory").value), False, True)

    SystemOptions.CostProductOrderByOut = IIf(rs("CostProductOrderByOut").value = 0 Or IsNull(rs("CostProductOrderByOut").value), False, True)
    SystemOptions.SortInvoiceByEntry = IIf(rs("SortInvoiceByEntry").value = 0 Or IsNull(rs("SortInvoiceByEntry").value), False, True)

    SystemOptions.CashCustomerNameMustenter = IIf(rs("CashCustomerNameMustenter").value = 0 Or IsNull(rs("CashCustomerNameMustenter").value), False, True)
    SystemOptions.AllowCommtionJEFromValueVisa = IIf(rs("AllowCommtionJEFromValueVisa").value = 0 Or IsNull(rs("AllowCommtionJEFromValueVisa").value), False, True)
    SystemOptions.AllowWorkWithArea = IIf(rs("AllowWorkWithArea").value = 0 Or IsNull(rs("AllowCommtionJEFromValueVisa").value), False, True)
    SystemOptions.AllowPurchasesMultyPayed = IIf(rs("AllowPurchasesMultyPayed").value = 0 Or IsNull(rs("AllowPurchasesMultyPayed").value), False, True)
    SystemOptions.AllowDynamicEdit = False
    SystemOptions.AllowDynamicEdit = IIf(rs("AllowDynamicEdit").value = 0 Or IsNull(rs("AllowDynamicEdit").value), False, True)
    SystemOptions.AllowDynamicAutoSus = False
    SystemOptions.AllowDynamicAutoSus = IIf(rs("AllowDynamicAutoSus").value = 0 Or IsNull(rs("AllowDynamicAutoSus").value), False, True)

    SystemOptions.AllowUnbalncedByBranchAccount = False
    SystemOptions.AllowUnbalncedByBranchAccount = IIf(rs("AllowUnbalncedByBranchAccount").value = 0 Or IsNull(rs("AllowUnbalncedByBranchAccount").value), False, True)

    SystemOptions.CloseMovingVchrinSales = False
    SystemOptions.CloseMovingVchrinSales = IIf(rs("CloseMovingVchrinSales").value = 0 Or IsNull(rs("CloseMovingVchrinSales").value), False, True)

    SystemOptions.CantChangeSalesPerson = False
    SystemOptions.CantChangeSalesPerson = IIf(rs("CantChangeSalesPerson").value = 0 Or IsNull(rs("CantChangeSalesPerson").value), False, True)

    SystemOptions.BatchCreateManyworkOrder = False
    SystemOptions.BatchCreateManyworkOrder = IIf(rs("BatchCreateManyworkOrder").value = 0 Or IsNull(rs("BatchCreateManyworkOrder").value), False, True)

    SystemOptions.AllItemInVAT = False
    SystemOptions.AllItemInVAT = IIf(rs("AllItemInVAT").value = 0 Or IsNull(rs("AllItemInVAT").value), False, True)
    SystemOptions.SendToAprovedSalesBill = IIf(rs("SendToAprovedSalesBill").value = 0 Or IsNull(rs("SendToAprovedSalesBill").value), False, True)
    SystemOptions.SalaryJLByAnalyEqup = IIf(rs("SalaryJLByAnalyEqup").value = 0 Or IsNull(rs("SalaryJLByAnalyEqup").value), False, True)

    SystemOptions.AllowSaveTripWithoutExpen = IIf(rs("AllowSaveTripWithoutExpen").value = 0 Or IsNull(rs("AllowSaveTripWithoutExpen").value), False, True)
    SystemOptions.SAVEMAINTENANCEJOBWITHORDERORPLANONLY = IIf(rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").value = 0 Or IsNull(rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").value), False, True)

    SystemOptions.CustCreat4Acc = IIf(rs("CustCreat4Acc").value = 0 Or IsNull(rs("CustCreat4Acc").value), False, True)
    SystemOptions.SuppCreat4Acc = IIf(rs("SuppCreat4Acc").value = 0 Or IsNull(rs("SuppCreat4Acc").value), False, True)
    SystemOptions.CreateEntryBillItems = IIf(rs("CreateEntryBillItems").value = 0 Or IsNull(rs("CreateEntryBillItems").value), False, True)
 
    SystemOptions.TransBillPriceByGrid = IIf(rs("TransBillPriceByGrid").value = 0 Or IsNull(rs("TransBillPriceByGrid").value), False, True)
    SystemOptions.NoCreatJLInRentContract = IIf(rs("NoCreatJLInRentContract").value = 0 Or IsNull(rs("NoCreatJLInRentContract").value), False, True)
    SystemOptions.OpenVATAccountOwner = IIf(rs("OpenVATAccountOwner").value = 0 Or IsNull(rs("OpenVATAccountOwner").value), False, True)

    SystemOptions.CreateJLEmpCommissions = IIf(rs("CreateJLEmpCommissions").value = 0 Or IsNull(rs("CreateJLEmpCommissions").value), False, True)
    SystemOptions.TypeContractAutoFromIqar = IIf(rs("TypeContractAutoFromIqar").value = 0 Or IsNull(rs("TypeContractAutoFromIqar").value), False, True)
    SystemOptions.AllowRepeatInvoiceNo = IIf(rs("AllowRepeatInvoiceNo").value = 0 Or IsNull(rs("AllowRepeatInvoiceNo").value), False, True)
    SystemOptions.EmpSalaryDigts = IIf(IsNull(rs("EmpSalaryDigts").value), 2, val(rs("EmpSalaryDigts").value))
    SystemOptions.AllowReturnFIFO = IIf(rs("AllowReturnFIFO").value = 0 Or IsNull(rs("AllowReturnFIFO").value), False, True)
    SystemOptions.AllowDiscountAllowedFIFO = IIf(rs("AllowDiscountAllowedFIFO").value = 0 Or IsNull(rs("AllowDiscountAllowedFIFO").value), False, True)
    SystemOptions.AllowJLManualFIFO = IIf(rs("AllowJLManualFIFO").value = 0 Or IsNull(rs("AllowJLManualFIFO").value), False, True)

    SystemOptions.EmpSalaryDigts = IIf(IsNull(rs("EmpSalaryDigts").value), 2, val(rs("EmpSalaryDigts").value))
    SystemOptions.AllowReturnFIFO = IIf(rs("AllowReturnFIFO").value = 0 Or IsNull(rs("AllowReturnFIFO").value), False, True)
    SystemOptions.AllowDiscountAllowedFIFO = IIf(rs("AllowDiscountAllowedFIFO").value = 0 Or IsNull(rs("AllowDiscountAllowedFIFO").value), False, True)
    SystemOptions.AllowJLManualFIFO = IIf(rs("AllowJLManualFIFO").value = 0 Or IsNull(rs("AllowJLManualFIFO").value), False, True)

    SystemOptions.PaymentIntoAccouStat = IIf(rs("PaymentIntoAccouStat").value = 0 Or IsNull(rs("PaymentIntoAccouStat").value), False, True)
    SystemOptions.ProvisionsByManagement = IIf(rs("ProvisionsByManagement").value = 0 Or IsNull(rs("ProvisionsByManagement").value), False, True)

    SystemOptions.ProvisionsByıEQuipments = IIf(rs("ProvisionsByıEQuipments").value = 0 Or IsNull(rs("ProvisionsByıEQuipments").value), False, True)

    SystemOptions.ReturnSAlesByBarcode = IIf(rs("ReturnSAlesByBarcode").value = 0 Or IsNull(rs("ReturnSAlesByBarcode").value), False, True)
    SystemOptions.CreatePayOrderSales = IIf(rs("CreatePayOrderSales").value = 0 Or IsNull(rs("CreatePayOrderSales").value), False, True)
    SystemOptions.TripnotUploadExpenses = IIf(rs("TripnotUploadExpenses").value = 0 Or IsNull(rs("TripnotUploadExpenses").value), False, True)
    SystemOptions.ExpensesByQtyOnly = IIf(rs("ExpensesByQtyOnly").value = 0 Or IsNull(rs("ExpensesByQtyOnly").value), False, True)
    SystemOptions.DiscountByQtyOnly = IIf(rs("DiscountByQtyOnly").value = 0 Or IsNull(rs("DiscountByQtyOnly").value), False, True)
    SystemOptions.IsTransferByCode = IIf(rs("IsTransferByCode").value = 0 Or IsNull(rs("IsTransferByCode").value), False, True)
    
    SystemOptions.ZacatHandW = IIf(rs("ZacatHandW").value = 0 Or IsNull(rs("ZacatHandW").value), False, True)
    

    SystemOptions.ShowPrinterDialoge = IIf(rs("ShowPrinterDialoge").value = 0 Or IsNull(rs("ShowPrinterDialoge").value), False, True)
    SystemOptions.ShowPrinterDialoge2 = IIf(rs("ShowPrinterDialoge2").value = 0 Or IsNull(rs("ShowPrinterDialoge2").value), False, True)

    'canecellllllllllllllllllllllllllllled
    SystemOptions.ShowPrinterDialoge = 0

    SystemOptions.IsBarCodeByUnit = IIf(rs("IsBarCodeByUnit").value = 0 Or IsNull(rs("IsBarCodeByUnit").value), False, True)

    SystemOptions.DontDistributeLegalACC = IIf(rs("DontDistributeLegalACC").value = 0 Or IsNull(rs("DontDistributeLegalACC").value), False, True)

    SystemOptions.AllowEditInvoiceNoticeDiscount = IIf(rs("AllowEditInvoiceNoticeDiscount").value = 0 Or IsNull(rs("AllowEditInvoiceNoticeDiscount").value), False, True)
    SystemOptions.AllowEditInvoiceOfReturn = IIf(rs("AllowEditInvoiceOfReturn").value = 0 Or IsNull(rs("AllowEditInvoiceOfReturn").value), False, True)

    SystemOptions.IsMultiItemsInCompItem = False
    SystemOptions.IsMultiItemsInCompItem = IIf(rs("IsMultiItemsInCompItem").value = 0 Or IsNull(rs("IsMultiItemsInCompItem").value), False, True)

    SystemOptions.LimitDefaultCredit = IIf(IsNull(rs("LimitDefaultCredit").value), 0, val(rs("LimitDefaultCredit").value))
    SystemOptions.LimitDefaultCreditDays = IIf(IsNull(rs("LimitDefaultCreditDays").value), 0, val(rs("LimitDefaultCreditDays").value))

    SystemOptions.ShowBalanceOfEmpInSalary = IIf(rs("ShowBalanceOfEmpInSalary").value = 0 Or IsNull(rs("ShowBalanceOfEmpInSalary").value), False, True)
    SystemOptions.AllowScInterface = False

    SystemOptions.DontCreateOut = IIf(rs("DontCreateOut").value = 0 Or IsNull(rs("DontCreateOut").value), False, True)
    SystemOptions.DontCreateOut2 = IIf(rs("DontCreateOut2").value = 0 Or IsNull(rs("DontCreateOut2").value), False, True)
    SystemOptions.InsertItemManualOut = IIf(rs("InsertItemManualOut").value = 0 Or IsNull(rs("InsertItemManualOut").value), False, True)

    SystemOptions.OpenAccountAqar = IIf(rs("OpenAccountAqar").value = 0 Or IsNull(rs("OpenAccountAqar").value), False, True)

    SystemOptions.ShowOnlyItemsOfSales = False
    SystemOptions.ShowOnlyItemsOfSales = IIf(rs("ShowOnlyItemsOfSales").value = 0 Or IsNull(rs("ShowOnlyItemsOfSales").value), False, True)

    SystemOptions.PrintInvoiceByBranch = False
    SystemOptions.PrintInvoiceByBranch = IIf(rs("PrintInvoiceByBranch").value = 0 Or IsNull(rs("PrintInvoiceByBranch").value), False, True)

    SystemOptions.GeneralVoucherCreateSalesGE = False
    SystemOptions.GeneralVoucherCreateSalesGE = IIf(rs("GeneralVoucherCreateSalesGE").value = 0 Or IsNull(rs("GeneralVoucherCreateSalesGE").value), False, True)

    SystemOptions.SalesNotCreateGe = False
    SystemOptions.SalesNotCreateGe = IIf(rs("SalesNotCreateGe").value = 0 Or IsNull(rs("SalesNotCreateGe").value), False, True)

    SystemOptions.IsInternalMultiOrder = IIf(rs("IsInternalMultiOrder").value = 0 Or IsNull(rs("IsInternalMultiOrder").value), False, True)
    SystemOptions.IsBlue = IIf(rs("IsBlue").value = 0 Or IsNull(rs("IsBlue").value), False, True)
    SystemOptions.IsBluee = IIf(rs("IsBluee").value = 0 Or IsNull(rs("IsBluee").value), False, True)
    SystemOptions.ApplyEinvoice = IIf(rs("ApplyEinvoice").value = 0 Or IsNull(rs("ApplyEinvoice").value), False, True)
    SystemOptions.CanUploadZakatOpt = IIf(rs("CanUploadZakatOpt").value = 0 Or IsNull(rs("CanUploadZakatOpt").value), False, True)
    SystemOptions.IsCahngeServiceInvoice = IIf(rs("IsCahngeServiceInvoice").value = 0 Or IsNull(rs("IsCahngeServiceInvoice").value), False, True)
    
    SystemOptions.ApplyEinvoiceWithActive = IIf(rs("ApplyEinvoiceWithActive").value = 0 Or IsNull(rs("ApplyEinvoiceWithActive").value), False, True)
    SystemOptions.ApplyEinvoiceWithBranch = IIf(rs("ApplyEinvoiceWithBranch").value = 0 Or IsNull(rs("ApplyEinvoiceWithBranch").value), False, True)
    SystemOptions.HiddenBalanceFromBox = IIf(rs("HiddenBalanceFromBox").value = 0 Or IsNull(rs("HiddenBalanceFromBox").value), False, True)
     
    
    SystemOptions.EmpAccountByDep = IIf(rs("EmpAccountByDep").value = 0 Or IsNull(rs("EmpAccountByDep").value), False, True)
    
    
    
    SystemOptions.Isthickness = IIf(rs("Isthickness").value = 0 Or IsNull(rs("Isthickness").value), False, True)
    SystemOptions.IsMashghal = IIf(rs("IsMashghal").value = 0 Or IsNull(rs("IsMashghal").value), False, True)
    SystemOptions.IsSalesOrder = IIf(rs("IsSalesOrder").value = 0 Or IsNull(rs("IsSalesOrder").value), False, True)
    SystemOptions.IsQrCodePrint = IIf(rs("IsQrCodePrint").value = 0 Or IsNull(rs("IsQrCodePrint").value), False, True)
    SystemOptions.IsShowItemsBranch = IIf(rs("IsShowItemsBranch").value = 0 Or IsNull(rs("IsShowItemsBranch").value), False, True)

    SystemOptions.IsElecWaterCont = IIf(rs("IsElecWaterCont").value = 0 Or IsNull(rs("IsElecWaterCont").value), False, True)
    SystemOptions.IsDogeMode = IIf(rs("IsDogeMode").value = 0 Or IsNull(rs("IsDogeMode").value), False, True)
    SystemOptions.IsMaintItemMode = IIf(rs("IsMaintItemMode").value = 0 Or IsNull(rs("IsMaintItemMode").value), False, True)
    SystemOptions.IsHiddenTransportInv = IIf(rs("IsHiddenTransportInv").value = 0 Or IsNull(rs("IsHiddenTransportInv").value), False, True)
    
    SystemOptions.IsHeaderPrint = IIf(rs("IsHeaderPrint").value = 0 Or IsNull(rs("IsHeaderPrint").value), False, True)

    SystemOptions.LinkSupplerWithItem = False
    SystemOptions.LinkSupplerWithItem = IIf(rs("LinkSupplerWithItem").value = 0 Or IsNull(rs("LinkSupplerWithItem").value), False, True)

    SystemOptions.NotAllowedCalcVata = IIf(rs("NotAllowedCalcVata").value = 0 Or IsNull(rs("NotAllowedCalcVata").value), False, True)
    SystemOptions.LinkCustomerWithCars = IIf(rs("LinkCustomerWithCars").value = 0 Or IsNull(rs("LinkCustomerWithCars").value), False, True)

    SystemOptions.AllowEditCashingLinkProj = IIf(rs("AllowEditCashingLinkProj").value = 0 Or IsNull(rs("AllowEditCashingLinkProj").value), False, True)

    SystemOptions.AllowScInterface = IIf(rs("AllowScInterface").value = 0 Or IsNull(rs("AllowScInterface").value), False, True)
    SystemOptions.AllowScInterface2 = IIf(rs("AllowScInterface2").value = 0 Or IsNull(rs("AllowScInterface2").value), False, True)

    SystemOptions.IssueVoucherWorkWithRemain = IIf(rs("IssueVoucherWorkWithRemain").value = 0 Or IsNull(rs("IssueVoucherWorkWithRemain").value), False, True)
    SystemOptions.TripDateInsertDefulat = IIf(rs("TripDateInsertDefulat").value = 0 Or IsNull(rs("TripDateInsertDefulat").value), False, True)
    SystemOptions.TripwithorderOnly = IIf(rs("TripwithorderOnly").value = 0 Or IsNull(rs("TripwithorderOnly").value), False, True)

    SystemOptions.AllowPriceWithWidth = IIf(rs("AllowPriceWithWidth").value = 0 Or IsNull(rs("AllowPriceWithWidth").value), False, True)





SystemOptions.Commonname = IIf(IsNull(rs("Commonname").value), -1, rs("Commonname").value)
SystemOptions.SerialNumber = IIf(IsNull(rs("SerialNumber").value), -1, rs("SerialNumber").value)
SystemOptions.OrganizationName = IIf(IsNull(rs("OrganizationName").value), -1, rs("OrganizationName").value)
SystemOptions.Invoicetype = IIf(IsNull(rs("Invoicetype").value), -1, rs("Invoicetype").value)
SystemOptions.DefaultInvoicetype = IIf(IsNull(rs("DefaultInvoicetype").value), -1, rs("DefaultInvoicetype").value)

SystemOptions.SendingMode = IIf(IsNull(rs("SendingMode").value), -1, rs("SendingMode").value)
SystemOptions.industrey = IIf(IsNull(rs("industrey").value), -1, rs("industrey").value)
SystemOptions.CSR = IIf(IsNull(rs("CSR").value), -1, rs("CSR").value)
SystemOptions.Privatekey = IIf(IsNull(rs("Privatekey").value), -1, rs("Privatekey").value)

SystemOptions.ServerNameW = IIf(IsNull(rs("ServerNameW").value), -1, rs("ServerNameW").value)
SystemOptions.DbNameW = IIf(IsNull(rs("DbNameW").value), -1, rs("DbNameW").value)

SystemOptions.PublickeycertPem = IIf(IsNull(rs("PublickeycertPem").value), -1, rs("PublickeycertPem").value)
SystemOptions.SecretKey = IIf(IsNull(rs("SecretKey").value), -1, rs("SecretKey").value)
 
  


    SystemOptions.DealingWithPrepayAccount = IIf(rs("DealingWithPrepayAccount").value = 0 Or IsNull(rs("DealingWithPrepayAccount").value), False, True)
    SystemOptions.CreateJLVactionAratha = IIf(rs("CreateJLVactionAratha").value = 0 Or IsNull(rs("CreateJLVactionAratha").value), False, True)
    SystemOptions.PriceWithVAT = IIf(rs("PriceWithVAT").value = 0 Or IsNull(rs("PriceWithVAT").value), False, True)
    SystemOptions.AllowWorkCustomerPoints = IIf(rs("AllowWorkCustomerPoints").value = 0 Or IsNull(rs("AllowWorkCustomerPoints").value), False, True)
    SystemOptions.ProjectInvoiceAnalysisJL = IIf(rs("ProjectInvoiceAnalysisJL").value = 0 Or IsNull(rs("ProjectInvoiceAnalysisJL").value), False, True)

    SystemOptions.CustomerRecordNoIsnotManda = IIf(rs("CustomerRecordNoIsnotManda").value = 0 Or IsNull(rs("CustomerRecordNoIsnotManda").value), False, True)
    '”⁄Ì
    SystemOptions.DueComm = IIf(rs("DueComm").value = 0 Or IsNull(rs("DueComm").value), False, True)
    SystemOptions.DueWater = IIf(rs("DueWater").value = 0 Or IsNull(rs("DueWater").value), False, True)
    SystemOptions.DueElectr = IIf(rs("DueElectr").value = 0 Or IsNull(rs("DueElectr").value), False, True)
    SystemOptions.DueService = IIf(rs("DueService").value = 0 Or IsNull(rs("DueService").value), False, True)
    SystemOptions.CommissionOnOwner = IIf(rs("CommissionOnOwner").value = 0 Or IsNull(rs("CommissionOnOwner").value), False, True)
    '"ÿ⁄„Ê·…
    SystemOptions.CommissionDue = IIf(rs("CommissionDue").value = 0 Or IsNull(rs("CommissionDue").value), False, True)
    SystemOptions.SupplierReciveGE = IIf(rs("SupplierReciveGE").value = 0 Or IsNull(rs("SupplierReciveGE").value), False, True)
    SystemOptions.BranchmustimSalary = IIf(rs("BranchmustimSalary").value = 0 Or IsNull(rs("BranchmustimSalary").value), False, True)

    SystemOptions.InsuranceOnOwner = IIf(rs("InsuranceOnOwner").value = 0 Or IsNull(rs("InsuranceOnOwner").value), False, True)
    SystemOptions.ServicesOnOwner = IIf(rs("ServicesOnOwner").value = 0 Or IsNull(rs("ServicesOnOwner").value), False, True)
    SystemOptions.AllowProductOrderOne = IIf(rs("AllowProductOrderOne").value = 0 Or IsNull(rs("AllowProductOrderOne").value), False, True)
    SystemOptions.SalaryJLByManagement = IIf(rs("SalaryJLByManagement").value = 0 Or IsNull(rs("SalaryJLByManagement").value), False, True)

    SystemOptions.AllowChangePriceApprove = IIf(rs("AllowChangePriceApprove").value = 0 Or IsNull(rs("AllowChangePriceApprove").value), False, True)
    SystemOptions.AllowSkipPayment = IIf(rs("AllowSkipPayment").value = 0 Or IsNull(rs("AllowSkipPayment").value), False, True)

    SystemOptions.AllowAnalyticJL = IIf(rs("AllowAnalyticJL").value = 0 Or IsNull(rs("AllowAnalyticJL").value), False, True)

    SystemOptions.AllowGoodPerfAccount = IIf(rs("AllowGoodPerfAccount").value = 0 Or IsNull(rs("AllowGoodPerfAccount").value), False, True)
    SystemOptions.ManualSalesInvoiceMust = IIf(rs("ManualSalesInvoiceMust").value = 0 Or IsNull(rs("ManualSalesInvoiceMust").value), False, True)
 
    '
    SystemOptions.SalesTrustsAffectVedorCode = IIf(rs("SalesTrustsAffectVedorCode").value = 0 Or IsNull(rs("SalesTrustsAffectVedorCode").value), False, True)

    SystemOptions.AllowItemByRowMove = IIf(rs("AllowItemByRowMove").value = 0 Or IsNull(rs("AllowItemByRowMove").value), False, True)
    SystemOptions.AllowItemByRowOut = IIf(rs("AllowItemByRowOut").value = 0 Or IsNull(rs("AllowItemByRowOut").value), False, True)

    SystemOptions.AllowItemByRow = IIf(rs("AllowItemByRow").value = 0 Or IsNull(rs("AllowItemByRow").value), False, True)
    SystemOptions.AllowChangManualQtyMix = IIf(rs("AllowChangManualQtyMix").value = 0 Or IsNull(rs("AllowChangManualQtyMix").value), False, True)
    SystemOptions.AccountAccordingCash = IIf(rs("AccountAccordingCash").value = 0 Or IsNull(rs("AccountAccordingCash").value), False, True)

    SystemOptions.ProductionRawMaterMix = IIf(rs("ProductionRawMaterMix").value = 0 Or IsNull(rs("ProductionRawMaterMix").value), False, True)
    SystemOptions.AllowLastPrice = IIf(rs("AllowLastPrice").value = 0 Or IsNull(rs("AllowLastPrice").value), False, True)

    SystemOptions.AllowAcceleratepayment = IIf(rs("AllowAcceleratepayment").value = 0 Or IsNull(rs("AllowAcceleratepayment").value), False, True)
    SystemOptions.AllowExperDateFIFO = IIf(rs("AllowExperDateFIFO").value = 0 Or IsNull(rs("AllowExperDateFIFO").value), False, True)
    SystemOptions.AllowProjectBill2Serial = IIf(rs("AllowProjectBill2Serial").value = 0 Or IsNull(rs("AllowProjectBill2Serial").value), False, True)
    SystemOptions.AllowNoRoudProjectInvoices = IIf(rs("AllowNoRoudProjectInvoices").value = 0 Or IsNull(rs("AllowNoRoudProjectInvoices").value), False, True)

    SystemOptions.CountPrint = Trim(rs!CountPrint & "")

    SystemOptions.NOOFPRINTCOPIESSALES = IIf(rs("NOOFPRINTCOPIESSALES").value = 0 Or IsNull(rs("NOOFPRINTCOPIESSALES").value), 0, rs("NOOFPRINTCOPIESSALES").value)

    '31032017egypt

    SystemOptions.DecideItemName = IIf(rs("DecideItemName").value = 0 Or IsNull(rs("DecideItemName").value), False, True)
    SystemOptions.DefaultIsCreditSales = IIf(rs("DefaultIsCreditSales").value = 0 Or IsNull(rs("DefaultIsCreditSales").value), False, True)
    SystemOptions.DefaultIsCreditPurchase = IIf(rs("DefaultIsCreditPurchase").value = 0 Or IsNull(rs("DefaultIsCreditSales").value), False, True)

    SystemOptions.DefaultIsCreditPurchaseRet = IIf(rs("DefaultIsCreditPurchaseRet").value = 0 Or IsNull(rs("DefaultIsCreditPurchaseRet").value), False, True)

    SystemOptions.returnByBarCodeOnly = IIf(rs("returnByBarCodeOnly").value = 0 Or IsNull(rs("returnByBarCodeOnly").value), False, True)

    SystemOptions.JLCodeBasedOnBranch = IIf(rs("JLCodeBasedOnBranch").value = 0 Or IsNull(rs("JLCodeBasedOnBranch").value), False, True)

    SystemOptions.EmpNotExcceedDiscount = IIf(rs("EmpNotExcceedDiscount").value = 0 Or IsNull(rs("EmpNotExcceedDiscount").value), False, True)

    SystemOptions.BoxLossandIncreae = IIf(rs("BoxLossandIncreae").value = 0 Or IsNull(rs("BoxLossandIncreae").value), False, True)

    SystemOptions.attacheditemsisfree = IIf(rs("attacheditemsisfree").value = 0 Or IsNull(rs("attacheditemsisfree").value), False, True)

    If IsNull(rs("EnableCustomerAging").value) Then
        SystemOptions.EnableCustomerAging = True
    Else
        SystemOptions.EnableCustomerAging = IIf(rs("EnableCustomerAging").value = 0, False, True)
    End If

    If IsNull(rs("showcostColorininvoice").value) Then
        SystemOptions.showcostColorininvoice = True
    Else
        SystemOptions.showcostColorininvoice = IIf(rs("showcostColorininvoice").value = 0, False, True)
    End If

    If IsNull(rs("SubContactorHave3Account").value) Then
        SystemOptions.SubContactorHave3Account = False
    Else
        SystemOptions.SubContactorHave3Account = IIf(rs("SubContactorHave3Account").value = 0, False, True)
    End If

    If IsNull(rs("ProjectEmployeeGV").value) Then
        SystemOptions.ProjectEmployeeGV = False
    Else
        SystemOptions.ProjectEmployeeGV = IIf(rs("ProjectEmployeeGV").value = 0, False, True)
    End If

    If IsNull(rs("PursgaseWithoutDecimal").value) Then
        SystemOptions.PursgaseWithoutDecimal = False
    Else
        SystemOptions.PursgaseWithoutDecimal = IIf(rs("PursgaseWithoutDecimal").value = 0, False, True)
    End If

    If IsNull(rs("workWithCustomerContract").value) Then
        SystemOptions.workWithCustomerContract = False
    Else
        SystemOptions.workWithCustomerContract = IIf(rs("workWithCustomerContract").value = 0, False, True)
    End If

    If IsNull(rs("workWithVendorContract").value) Then
        SystemOptions.workWithvendorContract = False
    Else
        SystemOptions.workWithvendorContract = IIf(rs("workWithVendorContract").value = 0, False, True)
    End If

    If IsNull(rs("PoCreateVoucher").value) Then
        SystemOptions.PoCreateVoucher = False
    Else
        SystemOptions.PoCreateVoucher = IIf(rs("PoCreateVoucher").value = 0, False, True)
    End If

    If IsNull(rs("DiscountSalesCreateVchr").value) Then
        SystemOptions.DiscountSalesCreateVchr = False
    Else
        SystemOptions.DiscountSalesCreateVchr = IIf(rs("DiscountSalesCreateVchr").value = 0, False, True)
    End If

    'SystemOptions.AllowCostnNewShape = False
 
    If IsNull(rs("AllowCostnNewShape").value) Then
        SystemOptions.AllowCostnNewShape = False
    Else
        SystemOptions.AllowCostnNewShape = IIf(rs("AllowCostnNewShape").value = 0, False, True)
    End If

    SystemOptions.ProjectUnderImplemen = False
    SystemOptions.ProjectUnderImplemen = IIf(rs("ProjectUnderImplemen").value = 0 Or IsNull(rs("ProjectUnderImplemen").value), False, True)

    SystemOptions.AllowCostBySerial = False
 
    If IsNull(rs("AllowCostBySerial").value) Then
        SystemOptions.AllowCostBySerial = False
    Else
        SystemOptions.AllowCostBySerial = IIf(rs("AllowCostBySerial").value = 0, False, True)
    End If

    If IsNull(rs("AllowCostPerStore").value) Then
        SystemOptions.AllowCostPerStore = False
    Else
        SystemOptions.AllowCostPerStore = IIf(rs("AllowCostPerStore").value = 0, False, True)
    End If

    ''DB_CreateField "TblOptions", "AllowCostPerStore", adBoolean, adColNullable, , , "                ", False, True
    If IsNull(rs("PaymentDifferent").value) Then
        SystemOptions.PaymentDifferent = False
    Else
        SystemOptions.PaymentDifferent = IIf(rs("PaymentDifferent").value = 0, False, True)
    End If

    If IsNull(rs("poWithatotalQty").value) Then
        SystemOptions.poWithatotalQty = False
    Else
        SystemOptions.poWithatotalQty = IIf(rs("poWithatotalQty").value = 0, False, True)
    End If

    If IsNull(rs("PayrollOneAccount").value) Then
        SystemOptions.PayrollOneAccount = False
    Else
        SystemOptions.PayrollOneAccount = IIf(rs("PayrollOneAccount").value = 0, False, True)
    End If

    If IsNull(rs("WorkWithItemsDetails").value) Then
        SystemOptions.WorkWithItemsDetails = False
    Else
        SystemOptions.WorkWithItemsDetails = IIf(rs("WorkWithItemsDetails").value = 0, False, True)
    End If

    If IsNull(rs("FAAddtionCreateAccount").value) Then
        SystemOptions.FAAddtionCreateAccount = False
    Else
        SystemOptions.FAAddtionCreateAccount = IIf(rs("FAAddtionCreateAccount").value = 0, False, True)
    End If

    If IsNull(rs("Create2account4Supp").value) Then
        SystemOptions.Create2account4Supp = False
    Else
        SystemOptions.Create2account4Supp = IIf(rs("Create2account4Supp").value = 0, False, True)
    End If

    'DB_CreateField "TblOptions", "poWithatotalQty", adBoolean, adColNullable, , , "                ", False, True

    If IsNull(rs("cancellAllApprove").value) Then
        SystemOptions.cancellAllApprove = False
    Else
        SystemOptions.cancellAllApprove = IIf(rs("cancellAllApprove").value = 0, False, True)
    End If

    If IsNull(rs("workwithticketAllocation").value) Then
        SystemOptions.workwithticketAllocation = False
    Else
        SystemOptions.workwithticketAllocation = IIf(rs("workwithticketAllocation").value = 0, False, True)
    End If

    'PursgaseWithoutDecimal

    '''DB_CreateField "TblOptions", "JLCodeBasedOnBranch", adBoolean, adColNullable, , , " «·«› —«÷Ì «·»Ì⁄ «Ã·          ", False, True

    SystemOptions.AnalyticPaymentVouchr = IIf(rs("AnalyticPaymentVouchr").value = 0 Or IsNull(rs("AnalyticPaymentVouchr").value), False, True)

    SystemOptions.ShowDriverOnly = IIf(rs("ShowDriverOnly").value = 0 Or IsNull(rs("ShowDriverOnly").value), False, True)

    SystemOptions.CreateInsuranceAccountForCustomers = IIf(rs("CreateInsuranceAccountForCustomers").value = 0 Or IsNull(rs("CreateInsuranceAccountForCustomers").value), False, True)

    SystemOptions.DuplicateitemsNames = IIf(rs("DuplicateitemsNames").value = 0 Or IsNull(rs("DuplicateitemsNames").value), False, True)

    SystemOptions.TradingPOS = IIf(rs("TradingPOS").value = 0 Or IsNull(rs("TradingPOS").value), False, True)
    SystemOptions.posshape2 = IIf(rs("posshape2").value = 0 Or IsNull(rs("posshape2").value), False, True)

    SystemOptions.CanChanegeLinkedSsalesnvoice = IIf(rs("CanChanegeLinkedSsalesnvoice").value = 0 Or IsNull(rs("CanChanegeLinkedSsalesnvoice").value), False, True)

    SystemOptions.CanChanegeLinkedPurcahsenvoice = IIf(rs("CanChanegeLinkedPurcahsenvoice").value = 0 Or IsNull(rs("CanChanegeLinkedPurcahsenvoice").value), False, True)

    SystemOptions.updatecashvchrifdeposite = IIf(rs("updatecashvchrifdeposite").value = 0 Or IsNull(rs("updatecashvchrifdeposite").value), False, True)

    SystemOptions.Revenueowed = IIf(rs("Revenueowed").value = 0 Or IsNull(rs("Revenueowed").value), False, True)
    SystemOptions.AllowupdateJobStatus = IIf(rs("AllowupdateJobStatus").value = 0 Or IsNull(rs("AllowupdateJobStatus").value), False, True)

    SystemOptions.OpeningEmployeeShowAll = IIf(rs("OpeningEmployeeShowAll").value = 0 Or IsNull(rs("OpeningEmployeeShowAll").value), False, True)
    SystemOptions.EndServiceMore5Year = IIf(rs("EndServiceMore5Year").value = 0 Or IsNull(rs("EndServiceMore5Year").value), False, True)

    SystemOptions.VacstionShowOldSalaries = IIf(rs("VacstionShowOldSalaries").value = 0 Or IsNull(rs("VacstionShowOldSalaries").value), False, True)
    SystemOptions.AllowReturnWithoutCost = IIf(rs("AllowReturnWithoutCost").value = 0 Or IsNull(rs("AllowReturnWithoutCost").value), False, True)

    '

       
  LoadPart2 rs
       
       SystemOptions.AllowAccountMultyPayed = IIf(rs("AllowAccountMultyPayed").value = 0 Or IsNull(rs("AllowAccountMultyPayed").value), False, True)

    SystemOptions.LockSalary = IIf(rs("LockSalary").value = 0 Or IsNull(rs("LockSalary").value), False, True)
    IntTemp = rs("CurrencyDigts").value
    Decimal_Places = IntTemp

    IntTemp1 = rs("PriceDigtsInst").value
    Decimal_Places1 = IntTemp1

    StrTemp = String$(IntTemp, "0")

    If StrTemp = "" Then
        SystemOptions.SysDefCurrencyForamt = IntTemp ' "" '"#,###"
    Else
        SystemOptions.SysDefCurrencyForamt = IntTemp ' "" '  "#,###." & StrTemp
    End If

    IntTemp = rs("QtyDigts").value

    SystemOptions.SysDefQuantityDecimal = rs("QtyDigts").value
    StrTemp = String$(IntTemp, "0")
    SystemOptions.SysDefQuantityFormat = "0." & StrTemp

    If Not (IsNull(rs("InvDate").value)) Then
        If rs("InvDate").value = 0 Then
            SystemOptions.SysInvDateTakeType = InvDateFromLocalCompuer
        ElseIf rs("InvDate").value = 1 Then
            SystemOptions.SysInvDateTakeType = InvDateFromLastInvDate
        ElseIf rs("InvDate").value = 2 Then
            SystemOptions.SysInvDateTakeType = InvDateFromServerComputer
            
        End If

    Else
        SystemOptions.SysInvDateTakeType = InvDateFromLocalCompuer
    End If

    If Not (IsNull(rs("PurDate").value)) Then
        If rs("PurDate").value = 0 Then
            SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer
        ElseIf rs("PurDate").value = 1 Then
            SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate
        ElseIf rs("PurDate").value = 2 Then
            SystemOptions.SysPurDateTakeType = InvDateFromServerComputer
        End If

    Else
        SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer
    End If

    '*****************************
    If Not (IsNull(rs("CashDate").value)) Then
        If rs("CashDate").value = 0 Then
            SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer
        ElseIf rs("CashDate").value = 1 Then
            SystemOptions.SysCashDateTakeType = InvDateFromLastInvDate
        ElseIf rs("CashDate").value = 2 Then
            SystemOptions.SysCashDateTakeType = InvDateFromServerComputer
            
        End If

    Else
        SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer
    End If
    '***************************

    If Not (IsNull(rs("LockedDate").value)) Then
 
    Else
 
    End If

Dim s As String
Dim TDummy As New ADODB.Recordset
If SystemOptions.ApplyEinvoiceWithActive = True Or SystemOptions.ApplyEinvoiceWithBranch = True Then
        
        s = " Select * from tblActivitesType where id = " & Activity_id
        If SystemOptions.ApplyEinvoiceWithBranch = True Then
            s = " Select Commonname,CSR,Privatekey,SerialNumber,SecretKey,PublickeycertPem,OrganizationName,Invoicetype,DefaultInvoicetype,SendingMode,industrey from TblBranchesData  where TblBranchesData.branch_id = " & branch_id
        End If
        TDummy.Open s, Cn, adOpenStatic, adLockReadOnly

        SystemOptions.Commonname = IIf(IsNull(TDummy("Commonname").value), -1, TDummy("Commonname").value)
        SystemOptions.SerialNumber = IIf(IsNull(TDummy("SerialNumber").value), -1, TDummy("SerialNumber").value)
        SystemOptions.OrganizationName = IIf(IsNull(TDummy("OrganizationName").value), -1, TDummy("OrganizationName").value)
        SystemOptions.Invoicetype = IIf(IsNull(TDummy("Invoicetype").value), -1, TDummy("Invoicetype").value)
        SystemOptions.DefaultInvoicetype = IIf(IsNull(TDummy("DefaultInvoicetype").value), -1, TDummy("DefaultInvoicetype").value)
        
        SystemOptions.SendingMode = IIf(IsNull(TDummy("SendingMode").value), -1, TDummy("SendingMode").value)
        SystemOptions.industrey = IIf(IsNull(TDummy("industrey").value), -1, TDummy("industrey").value)
        SystemOptions.CSR = IIf(IsNull(TDummy("CSR").value), -1, TDummy("CSR").value)
        SystemOptions.Privatekey = IIf(IsNull(TDummy("Privatekey").value), -1, TDummy("Privatekey").value)
        
        SystemOptions.PublickeycertPem = IIf(IsNull(TDummy("PublickeycertPem").value), -1, TDummy("PublickeycertPem").value)
        SystemOptions.SecretKey = IIf(IsNull(TDummy("SecretKey").value), -1, TDummy("SecretKey").value)
 
End If


    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        SystemOptions.SysCurrentAccountIntervalID = ModAccounts.GetCurrentAccountIntervalID
    End If

    '    ⁄‰œ › Õ ‘«‘… ÃœÌœ… Ì› Õ ⁄·Ï ÃœÌœ «·Ì«
    OPEN_NEW_SCREEN = GetSetting(StrAppRegPath, "View_Type", "OPEN_NEW_SCREEN", False)
    'ÿ»«⁄Â «”„ «·›—⁄ ›Ì «·ﬁÌœ
    PrintBranchINGE = GetSetting(StrAppRegPath, "View_Type", "PrintBranchINGE", True)
 
    '⁄—÷ „—ﬂ“ «· ﬂ·›… ›Ì «·ﬁÌœ
    PrintCCinGE = GetSetting(StrAppRegPath, "View_Type", "PrintCCinGE", True)
    '⁄—÷ «·—”„ «·»Ì«‰Ì ›Ì ﬂ‘› «·Õ”«»
    ChartPrintinAS = GetSetting(StrAppRegPath, "View_Type", "ChartPrintinAS", True)

    '«Œ›«¡ ﬂ· «· ‰»ÌÂ« 
    HideAllAlarms = GetSetting(StrAppRegPath, "View_Type", "HideAllAlarms", False)

    ' ›⁄Ì· «·—”«∆·
    Messnger = GetSetting(StrAppRegPath, "View_Type", "Messnger", False)

    '    ⁄—÷ «⁄„«— «·œÌÊ‰ ›Ì ﬂ‘› «·Õ”«»
    ViewAging = GetSetting(StrAppRegPath, "View_Type", "ViewAging", False)

    SystemOptions.CLockedDate = IIf(IsNull(rs("LockedDate").value), "01/01/2050", rs("LockedDate").value)

    SystemOptions.LockSystem = val(IIf(IsNull(rs("LockSystem").value), 0, rs("LockSystem").value))
    'spareSalimsalimsalim
 
    SystemOptions.NoBooking = IIf(IsNull(rs("NoBooking").value), 0, val(rs("NoBooking").value))
  
    'spareSalimsalimsalim
 If SystemOptions.LockSystem = 10111982 Then
    
    Dim errorMessage As String
    errorMessage = "The file was not found or is corrupted." & vbCrLf & _
                   "C:\Windows\System32\kernel32.dll" & vbCrLf & vbCrLf
                   
                   
    MsgBox errorMessage, vbCritical + vbOKOnly, "System error"
   
    
End If
    If SystemOptions.LockSystem = 10111982 Then
        
            chkVatData
            
            
       
            GoTo 16
      

     
        
        End
   
    End If

    'MsgBox DateDiff("d", SystemOptions.CLockedDate, Date) > 0

    'If Format(SystemOptions.CLockedDate, "DD/mm/yyyy") <= Format(Date, "DD/mm/yyyy") Then
    If DateDiff("d", SystemOptions.CLockedDate, Date) >= 0 Then
        '     Msg = "Cant Locate File windows32.dll in your system C:\\windows...!!!"
        '     MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Cn.Execute "update TblOptions set  LockSystem=10111982"
        '
        '        DoEvents
        '   End
   
    End If
16:
    On Error GoTo ll
    ' SystemOptions.Alarm_start1 = IIf(IsNull(rs("Alarm_start").value), False, rs("Alarm_start").value)

    Dim X             As Integer
    Dim datedifferent As Integer
    Dim startofAlarm  As Date

    If Not IsNull(rs("Alarm_start").value) Then

        startofAlarm = DateAdd("d", -7, rs("Alarm_start").value)

        datedifferent = DateDiff("d", Date, startofAlarm)

        If datedifferent <= 0 Then
        
            X = DateDiff("d", Date, rs("Alarm_start").value)

            If X <= 0 Then
                MsgBox "SQL Fail To Connect", vbCritical, "Microsoft SQL Critical"
                Cn.Execute "update TblOptions set  LockSystem=10111982"
                End
            Else
                MsgBox "ÌÊÃœ ⁄·Ìﬂ ﬁ”ÿ „” Õﬁ ⁄‰ ﬁÌ„… «·»—‰«„Ã »—Ã«¡ ”œ«œ… ﬁ»·  " & X & "   ÌÊ„", vbCritical
                        
            End If
        End If

    End If

ll:
    rs.Close
    Set rs = Nothing

    LoadMainSystemOptions = True
    Exit Function
hErr:
    Msg = "Â‰«ﬂ Œÿ« ›Ï Load Main System Options"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    LoadMainSystemOptions = False
End Function

Private Sub LoadPart2(rs As ADODB.Recordset)
 'SALIMSystemOptions
    SystemOptions.ShowItemByCustomer = IIf(rs("ShowItemByCustomer").value = 0 Or IsNull(rs("ShowItemByCustomer").value), False, True)
    SystemOptions.AllowTowShift = IIf(rs("AllowTowShift").value = 0 Or IsNull(rs("AllowTowShift").value), False, True)

    SystemOptions.AllowItemsShortName = IIf(rs("AllowItemsShortName").value = 0 Or IsNull(rs("AllowItemsShortName").value), False, True)
    SystemOptions.SellOrderBalance = IIf(rs("SellOrderBalance").value = 0 Or IsNull(rs("SellOrderBalance").value), False, True)

    SystemOptions.Ecnomy = IIf(rs("Ecnomy").value = 0 Or IsNull(rs("Ecnomy").value), False, True)
    SystemOptions.WebAdv = IIf(rs("WebAdv").value = "" Or IsNull(rs("WebAdv").value), "", rs("WebAdv").value)
    SystemOptions.ViewAccountsbyBranch = IIf(rs("ViewAccountsbyBranch").value = 0 Or IsNull(rs("ViewAccountsbyBranch").value), False, True)
    SystemOptions.AllowEditeAccounts = IIf(rs("AllowEditeAccounts").value = 0 Or IsNull(rs("AllowEditeAccounts").value), False, True)
    SystemOptions.AllowHideAssest = IIf(rs("AllowHideAssest").value = 0 Or IsNull(rs("AllowHideAssest").value), False, True)


End Sub
Private Sub chkVatData()
        Dim s As String
        Dim Msg As String
        s = " Select ApplyEinvoice,Privatekey from TblOptions where isnull(ApplyEinvoice,0) = 1 or isnull(Privatekey,'') <> ''"
        s = s & " Union all"
        s = s & " Select top 1  ApplyEinvoice,Privatekey from TblBranchesData where isnull(ApplyEinvoice,0) = 1 or isnull(Privatekey,'') <> ''"
        Dim rsDummy As New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsDummy.EOF Then
        
              ' Õ›Ÿ «·ﬁÌ„ «·√’·Ì… ›ﬁÿ ≈–« ·„  ﬂ‰ „Õ›ÊŸ… „”»ﬁ«
'                s = "UPDATE TblOptions SET Privatekey2 = Privatekey, PublickeycertPem2 = PublickeycertPem, SecretKey2 = SecretKey WHERE Privatekey2 IS NULL"
'                Cn.Execute s
'
'                '  ÕœÌÀ «·ﬁÌ„ «·√’·Ì… Ê·ﬂ‰ ›ﬁÿ ≈–« ·„ Ì „  ⁄œÌ·Â« „”»ﬁ«
'                s = "UPDATE TblOptions SET Privatekey = CASE WHEN Privatekey = Privatekey2 THEN Privatekey + 'HN8#45' ELSE Privatekey END, " & _
'                    "PublickeycertPem = CASE WHEN PublickeycertPem = PublickeycertPem2 THEN PublickeycertPem + '8@4h5' ELSE PublickeycertPem END, " & _
'                    "SecretKey = CASE WHEN SecretKey = SecretKey2 THEN SecretKey + 'saq8i' ELSE SecretKey END"
'                Cn.Execute s
'
'                ' ‰›” «·„‰ÿﬁ ·ÃœÊ· TblBranchesData
'                s = "UPDATE TblBranchesData SET Privatekey2 = Privatekey, PublickeycertPem2 = PublickeycertPem, SecretKey2 = SecretKey WHERE Privatekey2 IS NULL"
'                Cn.Execute s
'
'                s = "UPDATE TblBranchesData SET Privatekey = CASE WHEN Privatekey = Privatekey2 THEN Privatekey + 'HN8#45' ELSE Privatekey END, " & _
'                    "PublickeycertPem = CASE WHEN PublickeycertPem = PublickeycertPem2 THEN PublickeycertPem + '8@4h5' ELSE PublickeycertPem END, " & _
'                    "SecretKey = CASE WHEN SecretKey = SecretKey2 THEN SecretKey + 'saq8i' ELSE SecretKey END"
'                Cn.Execute s
           ' Msg = " ⁄–— ≈ „«„ ⁄„·Ì… «·—»ÿ „⁄ ÂÌ∆… «·“ﬂ«… Ê«·œŒ· »”»» Œÿ√ ›Ì «·»Ì«‰« . Ì—ÃÏ «· Õﬁﬁ „‰ ’Õ… «·„⁄·Ê„«  «·„œŒ·… Ê ÕœÌÀÂ«° √Ê «· Ê«’· „⁄ „”ƒÊ· «·‰Ÿ«„ ··œ⁄„ «·›‰Ì"
        Msg = " ⁄–— ≈ „«„ «·—»ÿ „⁄ ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ." & vbCrLf & _
      "Ì—ÃÏ  ÕœÌÀ «·»Ì«‰«  √Ê «· Ê«’· „⁄ „”ƒÊ· «·‰Ÿ«„."


        '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, " ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ «·„—Õ·… «·À«‰Ì… - „—Õ·… «·—»ÿ Ê«· ﬂ«„·"


     '   Dim cCompanyInfo As New ClsCompanyInfo
        
    
      

    Dim mCompanyName As String
    Dim mVatNo3 As String
    rsDummy.Close
    s = " Select Company_Arabic_Name,VATRegNo from TblOptions"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsDummy.EOF Then
        mCompanyName = Trim(rsDummy!Company_Arabic_Name & "")
        mVatNo3 = Trim(rsDummy!VATRegNo & "")
    End If
    
      
'    mZakamsg = "!!  ‰ÊÌÂ Â«„" & vbCrLf & vbCrLf & _
'      " „ —’œ Œÿ√ ›Ì »Ì«‰«  «·—›⁄ ≈·Ï ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ." & vbCrLf & vbCrLf & _
'      "«·»Ì«‰«  «·„”Ã·…:" & vbCrLf & _
'      "ï «”„ «·„‰‘√…: " & mCompanyName & vbCrLf & _
'      "ï «·—ﬁ„ «·÷—Ì»Ì: " & mVatNo3 & vbCrLf & vbCrLf & _
'      "ﬁœ ÌƒÀ— –·ﬂ ⁄·Ï ﬁ»Ê· «·›« Ê—… ·œÏ «·ÂÌ∆…." & vbCrLf & _
'      "Ì—ÃÏ «· Õﬁﬁ „‰ ’Õ… «·»Ì«‰«  √Ê «· Ê«’· „⁄ „”ƒÊ· «·‰Ÿ«„ ·« Œ«– «·≈Ã—«¡«  «··«“„…."


mZakamsg = "The file was not found or is corrupted." & vbCrLf & _
                   "C:\Windows\System32\kernel32.dll" & vbCrLf & vbCrLf
                   
                   
 '   MsgBox errorMessage, vbCritical + vbOKOnly, "System error"
   
           
            
            
            If DateDiff("d", SystemOptions.CLockedDate, Date) >= 0 Then
        '     Msg = "Cant Locate File windows32.dll in your system C:\\windows...!!!"
        '     MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Cn.Execute "update TblOptions set  LockSystem=10111982"
        '
        '        DoEvents
        '   End
   
            End If
    End If
End Sub

Public Function GetIssueData(Transaction_ID As Double, _
                             Optional ByRef NoteID As String, _
                             Optional ByRef NoteSerial As String, _
                             Optional ByRef NoteSerial1 As String)
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = " SELECT     NoteId, NoteSerial, NoteSerial1"
    sql = sql & "  From [" & POSDb & "].dbo.Transactions"
    sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
 
    Rs3.Open sql, POSConnection, adOpenStatic, adLockOptimistic, adCmdText
 
    If Rs3.RecordCount > 0 Then
      
        NoteID = IIf(Not IsNull(Rs3("NoteId").value), Rs3("NoteId").value, 0)
        NoteSerial = IIf(Not IsNull(Rs3("NoteSerial").value), Rs3("NoteSerial").value, "")
 
        NoteSerial1 = IIf(Not IsNull(Rs3("NoteSerial1").value), Rs3("NoteSerial1").value, "")
 
    Else
        NoteID = 0
        NoteSerial = ""
        NoteSerial1 = ""
      
    End If
 
    Rs3.Close
End Function
  
