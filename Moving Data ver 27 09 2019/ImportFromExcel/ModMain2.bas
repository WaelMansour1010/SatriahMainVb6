Attribute VB_Name = "ModMain2"
  
Option Explicit

Public SystemOptions As MainOptions

Public Const SW_SHOWNORMAL = 1
Public HidLowering As Boolean
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
    Public CurrentVersion As String
    
'this function to get handle for DeskTop
Public Declare Function GetDesktopWindow _
               Lib "user32" () As Long

'to find any window
Public Declare Function FindWindow _
               Lib "user32" _
               Alias "FindWindowA" (ByVal lpClassName As String, _
                                    ByVal lpWindowName As String) As Long

Public Declare Function GetWindowRect _
               Lib "user32" (ByVal hwnd As Long, _
                             lpRect As RECT) As Long

Public Declare Function GetWindowLong _
               Lib "user32" _
               Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong _
               Lib "user32" _
               Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long

Public Declare Function GetClientRect _
               Lib "user32" (ByVal hwnd As Long, _
                             lpRect As RECT) As Long

Public Declare Function InvalidateRect _
               Lib "user32" (ByVal hwnd As Long, _
                             lpRect As RECT, _
                             ByVal bErase As Long) As Long

Public Declare Function ClientToScreen _
               Lib "user32" (ByVal hwnd As Long, _
                             lpPoint As POINTAPI) As Long

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
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
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long)

'This API to show a form
Public Declare Function ShowWindow _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal nCmdShow As Long) As Long

Public Const SW_SHOW As Long = 5

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
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByRef pti As TITLEBARINFO) As Long

'this api used in the MouseDown_Event to enable the user From
'Move the form From any Postion ...
'Try this with MDI Form by Move it From any Position by the Mouse...
Public Declare Function ReleaseCapture _
               Lib "user32" () As Long

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
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
    
    SysRunNumber As Integer

    SysCurrentAccountIntervalID As Long  'ﬂÊœ «·› —… «·„Õ«”»Ì… «·Õ«·Ì…
    SysInvDateTakeType As TakeDateTypes
    SysPurDateTakeType As TakeDateTypes
    SysCashDateTakeType As TakeDateTypes
    
    SysMantainceAllow As Boolean
    SysAllowStockNegative As Boolean '«·”Õ» ⁄·Ï «·„ﬂ‘Ê› „‰ «·„Œ“‰
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
    ShowOnlyItemsOfSales As Boolean
    GeneralVoucherCreateSalesGE As Boolean
    SalesNotCreateGe As Boolean
PrintInvoiceByBranch As Boolean
     LinkSupplerWithItem As Boolean
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
ShowPrinterDialoge As Boolean

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




    DealingWithPrepayAccount As Boolean
NotAllowedCalcVata As Boolean
AllowSkipDiscountGroup As Boolean
OpenAtProduction As Boolean

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

     '31032017egypt
     'modmod
         CanCustomerandVendor As Boolean
         

CanEditOnlyPayMethod As Boolean

SortInvoiceByEntry As Boolean
CostProductOrderByOut As Boolean

             CanPartialpayment As Boolean
             EndRentifPayed As Boolean
             cantCahngeAkarinExpenses  As Boolean
             EmployeeSalaryBYBranch  As Boolean
             returnnotcreatvoucher As Boolean
             WaiverSetByContract As Boolean
NoBooking As Integer
MultyStore As Boolean
     RawMaterMix2 As Boolean 'modmod
              DontCreateOut As Boolean 'modmod
     DontCreateOut2 As Boolean 'modmod
     InsertItemManualOut As Boolean 'modmod
     
     
    UserInvoiceShowProfit As Integer
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
CanEditCars As Boolean

CanChangeOut As Boolean
CanCancelContract As Boolean

 AllowSaveTripWithoutExpen As Boolean
 SAVEMAINTENANCEJOBWITHORDERORPLANONLY As Boolean
 
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
    CustomerhavethreeAccounts As Boolean
    CustomerhavethreeAccounts1 As Boolean
    CostByProduction As Boolean
AllowRepeateCar As Boolean
  CanPrintMultiSales As Boolean
  CanPayWithoutPrint As Boolean
  MaintOrderCantRepeatSales As Boolean
MaintOrderCantRepeatBillBuy As Boolean
TripRevenueAuto As Boolean
IsByNewCoding As Boolean
IsAutoNameItems As Boolean


cdoSMTPServer  As String
TxtFromName  As String
txtFromEmail  As String

cdoSendUserName  As String
cdoSendPassword  As String
cdoSMTPUseSSL As Boolean

cdoSMTPServerPort As Integer

PaymentMethLaterCompItem As Boolean
ShowBalanceCustInv As Boolean

IsBreaks As Boolean
IsCodeByBranch As Boolean


    IsSerialByUserTrans As Boolean
   IsSerialByUserVouch As Boolean
    NoOFDigitUserTrans As Integer
 
  
  Breaks As String
  
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
    StoreAccountHaveSettelment As Boolean
    eachStoreHaveLossAccount As Boolean
    eachStoreHaveGiftAccount As Boolean
    AllowExperDateFIFO As Boolean
End Type

'Public FrmNewsBarPane As FrmPane '‘—Ìÿ «·√Œ»«— Ê«·„⁄·Ê„« 


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

Public Decimal_Places As Integer

Public Decimal_Places1 As Integer

Public PrintBranchINGE As Boolean

Public PrintCCinGE As Boolean

Public ChartPrintinAS As Boolean

Public HideAllAlarms As Boolean

Public Messnger As Boolean
 Public AlarmAuto As Boolean
Public ViewAging As Boolean
Dim sql As String
Public Declare Function InternetGetConnectedState Lib _
    "wininet" (ByRef dwFlags As Long, ByVal dwReserved As _
    Long) As Long
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
'
'Private Function CreateKey() As String
'    Dim Msg As String
'    Msg = GetHardDiskData(False)
'    Msg = Msg & "**" '& GetProcessorData(False)
'
'    Dim i As Integer
'    Dim StrOutKey As String
'    Dim StrChar As String
'    Dim LngOutKey As Long
'    Dim VarTemp As Variant
'    VarTemp = Split(Msg, "**", , vbBinaryCompare)
'
'    For i = 1 To Len(Msg)
'        StrChar = Mid(Msg, i, 1)
'
'        If StrChar <> "" Then
'            LngOutKey = LngOutKey + (Asc(StrChar) * 12)
'        End If
'
'    Next i
'
'    LngOutKey = (LngOutKey + LngOutKey)
'    LngOutKey = LngOutKey * 3
'    CreateKey = CStr(LngOutKey)
'End Function

Public Function GetMsgs(IntCode As Integer, _
                        IntButtons As VBA.VbMsgBoxStyle) As VBA.VbMsgBoxResult
    Dim Msg As String
    Dim BoxRtl As Long
    Dim StrPath As String
    Dim IntFile As Integer
    Dim MSGType As MsgData
    StrPath = App.Path
    StrPath = IIf(Right(StrPath, 1) = "\", "", StrPath & "\")
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


Private Sub CheckSerial()
    'On Error GoTo ErrTrap
    'Dim Strsql As String
    'Dim RsTemp As ADODB.Recordset
    'Strsql = "select RunCount form TblOptions "
    'Set RsTemp = New ADODB.Recordset
    'RsTemp.Open "TblOptions", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    'If Not (RsTemp.EOF Or RsTemp.BOF) Then
    '    If Val(RsTemp("RunCount")) <= 0 Then
    '        FrmCheckSerial.Show vbModal
    '    End If
    'End If
    'Exit Sub
    'If SystemOptions.SysRunNumber = 0 Then
    '    FrmCheckSerial.Show vbModal
    'End If
ErrTrap:
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

Public Sub ShowFormtitles(Frm As Form)
  Dim My_SQL As String
 Dim Msg As String
 Dim filename As String
 filename = App.Path & "\titles\" & Frm.Name & ".txt"
 
    If Dir(filename, vbNormal) = "" Then
            Msg = " File No found  ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
            
           Exit Sub
        End If
        
    Open filename For Input As #1
 
Dim a As String
Dim VarSet() As String
Dim Label() As String
Dim controlname As String
Dim controlindex As String
Dim ControlCaptionA  As String
Dim ControlCaptionE  As String
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
           
 
    Dim Txt As TextBox
    Dim lbl As Label
    Dim xTab As C1Tab
    Dim temp As C1Tab
    Dim ctl As Control
    Dim ConCtl  As Control
    Dim i  As Integer
    Dim ss As String
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
  Dim My_SQL As String
 Dim Msg As String
 Dim filename As String
 filename = App.Path & "\titles\" & Frm.Name & ".txt"
 
    If Dir(filename, vbNormal) = "" Then
            Msg = " File No found  ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
            
           Exit Sub
        End If
        
    Open filename For Input As #1
 
Dim a As String
Dim VarSet() As String
Dim Label() As String
Dim controlname As String
Dim controlindex As String
Dim ControlCaptionA  As String
Dim ControlCaptionE  As String
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
                           controlindex = Val(Label(1))
                           Else
                           controlindex = -1
                        End If
            
              End If
           
 
    Dim Txt As TextBox
    Dim lbl As Label
    Dim xTab As C1Tab
    Dim temp As C1Tab
    Dim ctl As Control
    Dim ConCtl  As Control
    Dim i  As Integer
    Dim ss As String
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

 
Public Function GetMyTitleBarHight(FrmHwnd As Long) As Single
    Dim TitleInfo As TITLEBARINFO
    Dim SngTemp As Single

    'Initialize structure
    TitleInfo.cbSize = Len(TitleInfo)
    'Retrieve information about the tilte bar of this window
    GetTitleBarInfo FrmHwnd, TitleInfo
    'Show some of that information
    SngTemp = TitleInfo.rcTitleBar.Bottom '- TitleInfo.rcTitleBar.Top
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

Public Function CorrectCurrency(SngNumber As Single) As Currency
    Dim IntIntgerFactor As Integer
    Dim IntFractional As Single
    Dim SngMod As Single

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

Public Function LoadMainSystemOptions2() As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim IntTemp As Integer
    Dim IntTemp1 As Integer

    Dim StrTemp As String
    Dim StrSQL  As String

    'On Error GoTo hErr
    On Error Resume Next
    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    SystemOptions.SysAllowStockNegative = IIf(rs("AllowStockNegative").Value = 0 Or IsNull(rs("AllowStockNegative").Value), False, True)
    SystemOptions.SysAllowBoxNegative = IIf(rs("AllowBoxNegative").Value = 0 Or IsNull(rs("AllowBoxNegative").Value), False, True)
    SystemOptions.SysMantainceAllow = IIf(rs("MantainceAllow").Value = 0 Or IsNull(rs("MantainceAllow").Value), False, True)
    
    SystemOptions.SysMainStockCostMethod = rs("MainStockCostType").Value
    SystemOptions.itemSeprator = rs("itemSeprator").Value
    ''''
    SystemOptions.ChasingStatus = IIf(rs("ChasingStatus").Value = 0 Or IsNull(rs("ChasingStatus").Value), 1, rs("ChasingStatus").Value)
    ''''
    
    
    SystemOptions.Items_or_operation = IIf(IsNull(rs("Items_or_operation").Value), -1, Val(rs("Items_or_operation").Value))
    
    SystemOptions.ProjectDiscountPolicy = IIf(IsNull(rs("ProjectDiscountPolicy").Value), 0, (rs("ProjectDiscountPolicy").Value))
    
  'ProjectDiscountPolicy
   
  
    SystemOptions.gldetails_or_gl_general = IIf(IsNull(rs("gl_detaila_or_total").Value), -1, Val(rs("gl_detaila_or_total").Value))
    SystemOptions.ProcessPeriodType = IIf(IsNull(rs("ProcessPeriodType").Value), -1, Val(rs("ProcessPeriodType").Value))

    SystemOptions.itemcodePart1 = IIf(IsNull(rs("itemcodePart1").Value), -1, Val(rs("itemcodePart1").Value))
    SystemOptions.itemcodePart2 = IIf(IsNull(rs("itemcodePart2").Value), -1, Val(rs("itemcodePart2").Value))
    SystemOptions.itemcodePart3 = IIf(IsNull(rs("itemcodePart3").Value), -1, Val(rs("itemcodePart3").Value))
    SystemOptions.itemcodePart1NoOFDigit = IIf(IsNull(rs("itemcodePart1NoOFDigit").Value), 0, Val(rs("itemcodePart1NoOFDigit").Value))
    SystemOptions.itemcodePart2NoOFDigit = IIf(IsNull(rs("itemcodePart2NoOFDigit").Value), 0, Val(rs("itemcodePart2NoOFDigit").Value))
    SystemOptions.itemcodePart3NoOFDigit = IIf(IsNull(rs("itemcodePart3NoOFDigit").Value), 0, Val(rs("itemcodePart3NoOFDigit").Value))
    SystemOptions.itemcodeSeperator1 = IIf(IsNull(rs("itemcodeSeperator1").Value), "", (rs("itemcodeSeperator1").Value))
    SystemOptions.itemcodeSeperator2 = IIf(IsNull(rs("itemcodeSeperator2").Value), "", (rs("itemcodeSeperator2").Value))
    SystemOptions.itemsWorkWithSize = IIf(rs("itemsWorkWithSize").Value = 0 Or IsNull(rs("itemsWorkWithSize").Value), False, True)
    
    
  
    SystemOptions.itemsWorkWithColor = IIf(rs("itemsWorkWithColor").Value = 0 Or IsNull(rs("itemsWorkWithColor").Value), False, True)
    SystemOptions.itemsWorkWithDates = IIf(rs("itemsWorkWithDates").Value = 0 Or IsNull(rs("itemsWorkWithDates").Value), False, True)
    SystemOptions.itemsWorkWithClass = IIf(rs("itemsWorkWithClass").Value = 0 Or IsNull(rs("itemsWorkWithClass").Value), False, True)


    SystemOptions.ItemcodeGroupOnly = IIf(rs("ItemcodeGroupOnly").Value = 0 Or IsNull(rs("ItemcodeGroupOnly").Value), False, True)
    SystemOptions.ItemcodeGroupandParentGroup = IIf(rs("ItemcodeGroupandParentGroup").Value = 0 Or IsNull(rs("ItemcodeGroupandParentGroup").Value), False, True)
 
    SystemOptions.SaleDiscount1 = IIf(IsNull(rs("SaleDiscount1").Value), 0, Val(rs("SaleDiscount1").Value))
    SystemOptions.SaleDiscount2 = IIf(IsNull(rs("SaleDiscount2").Value), 0, Val(rs("SaleDiscount2").Value))
    SystemOptions.SaleDiscount3 = IIf(IsNull(rs("SaleDiscount3").Value), 0, Val(rs("SaleDiscount3").Value))
    SystemOptions.SaleDiscount4 = IIf(IsNull(rs("SaleDiscount4").Value), 0, Val(rs("SaleDiscount4").Value))
    SystemOptions.autoIssueVoucher = IIf(rs("autoIssueVoucher").Value = 0 Or IsNull(rs("autoIssueVoucher").Value), False, True)
 
    SystemOptions.MonthIs30days = IIf(rs("MonthIs30days").Value = 0 Or IsNull(rs("MonthIs30days").Value), False, True)
 SystemOptions.RawMaterMix = IIf(rs("RawMaterMix").Value = 0 Or IsNull(rs("RawMaterMix").Value), False, True)
    SystemOptions.bankComm = IIf(rs("bankComm").Value = 0 Or IsNull(rs("bankComm").Value), False, True)
    SystemOptions.ChequeBox = IIf(rs("ChequeBox").Value = 0 Or IsNull(rs("ChequeBox").Value), False, True)
    SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").Value = 0 Or IsNull(rs("CustomerhavethreeAccounts").Value), False, True)
 SystemOptions.CustomerhavethreeAccounts1 = IIf(rs("CustomerhavethreeAccounts1").Value = 0 Or IsNull(rs("CustomerhavethreeAccounts1").Value), False, True)
 
  
  SystemOptions.AllowRepeateCar = IIf(rs("AllowRepeateCar").Value = 0 Or IsNull(rs("AllowRepeateCar").Value), False, True)
  SystemOptions.CostByProduction = IIf(rs("CostByProduction").Value = 0 Or IsNull(rs("CostByProduction").Value), False, True)


  'modmod
  SystemOptions.PaymentMethLaterCompItem = IIf(rs("PaymentMethLaterCompItem").Value = 0 Or IsNull(rs("PaymentMethLaterCompItem").Value), False, True)
  SystemOptions.ShowBalanceCustInv = IIf(rs("ShowBalanceCustInv").Value = 0 Or IsNull(rs("ShowBalanceCustInv").Value), False, True)
  
  
    SystemOptions.IsCodeByBranch = IIf(rs("IsCodeByBranch").Value = 0 Or IsNull(rs("IsCodeByBranch").Value), False, True)
      SystemOptions.IsBreaks = IIf(rs("IsBreaks").Value = 0 Or IsNull(rs("IsBreaks").Value), False, True)
  

  
  SystemOptions.MaintOrderCantRepeatSales = IIf(rs("MaintOrderCantRepeatSales").Value = 0 Or IsNull(rs("MaintOrderCantRepeatSales").Value), False, True)
SystemOptions.MaintOrderCantRepeatBillBuy = IIf(rs("MaintOrderCantRepeatBillBuy").Value = 0 Or IsNull(rs("MaintOrderCantRepeatBillBuy").Value), False, True)

SystemOptions.TripRevenueAuto = IIf(rs("TripRevenueAuto").Value = 0 Or IsNull(rs("TripRevenueAuto").Value), False, True)

SystemOptions.IsByNewCoding = IIf(rs("IsByNewCoding").Value = 0 Or IsNull(rs("IsByNewCoding").Value), False, True)
SystemOptions.IsAutoNameItems = IIf(rs("IsAutoNameItems").Value = 0 Or IsNull(rs("IsAutoNameItems").Value), False, True)

SystemOptions.cdoSMTPUseSSL = IIf(rs("cdoSMTPUseSSL").Value = 0 Or IsNull(rs("cdoSMTPUseSSL").Value), False, True)

SystemOptions.cdoSMTPServerPort = IIf(rs("cdoSMTPServerPort").Value = 0 Or IsNull(rs("cdoSMTPServerPort").Value), 587, rs("cdoSMTPServerPort").Value)

SystemOptions.cdoSMTPServer = IIf(rs("cdoSMTPServer").Value = "" Or IsNull(rs("cdoSMTPServer").Value), "cdoSMTPServer", rs("cdoSMTPServer").Value)

SystemOptions.cdoSendUserName = IIf(rs("cdoSendUserName").Value = "" Or IsNull(rs("cdoSendUserName").Value), "cdoSendUserName", rs("cdoSendUserName").Value)
SystemOptions.cdoSendPassword = IIf(rs("cdoSendPassword").Value = "" Or IsNull(rs("cdoSendPassword").Value), "cdoSendPassword", rs("cdoSendPassword").Value)

SystemOptions.TxtFromName = IIf(rs("txtFromName").Value = "" Or IsNull(rs("txtFromName").Value), " ", rs("txtFromName").Value)
SystemOptions.txtFromEmail = IIf(rs("txtFromEmail").Value = "" Or IsNull(rs("txtFromEmail").Value), " ", rs("txtFromEmail").Value)


 
 



SystemOptions.NoOFDigitUserTrans = IIf(Val(rs!NoOFDigitUserTrans & "") = 0, 2, Val(rs!NoOFDigitUserTrans & ""))
SystemOptions.NoOFDigitUserVouc = IIf(Val(rs!NoOFDigitUserVouc & "") = 0, 2, Val(rs!NoOFDigitUserVouc & ""))

SystemOptions.Breaks = IIf(Trim(rs!Breaks & "") = "", "", Trim(rs!Breaks & ""))



SystemOptions.IsSerialByUserTrans = IIf(rs("IsSerialByUserTrans").Value = 0 Or IsNull(rs("IsSerialByUserTrans").Value), False, True)
    SystemOptions.IsSerialByUserVouch = IIf(rs("IsSerialByUserVouch").Value = 0 Or IsNull(rs("IsSerialByUserVouch").Value), False, True)



  SystemOptions.IsSomeItemWeight = IIf(rs("IsSomeItemWeight").Value = 0 Or IsNull(rs("IsSomeItemWeight").Value), False, True)
  SystemOptions.IsMergeVat = IIf(rs("IsMergeVat").Value = 0 Or IsNull(rs("IsMergeVat").Value), False, True)
  
  SystemOptions.FromNo = IIf(IsNull(rs("FromNo").Value), 0, Val(rs("FromNo").Value))
SystemOptions.OrNo = IIf(IsNull(rs("OrNo").Value), 0, Val(rs("OrNo").Value))
SystemOptions.CodeFrom = IIf(IsNull(rs("CodeFrom").Value), 0, Val(rs("CodeFrom").Value))
SystemOptions.CodeTo = IIf(IsNull(rs("CodeTo").Value), 0, Val(rs("CodeTo").Value))
SystemOptions.WeightFrom = IIf(IsNull(rs("WeightFrom").Value), 0, Val(rs("WeightFrom").Value))
   SystemOptions.WeightTo = IIf(IsNull(rs("WeightTo").Value), 0, Val(rs("WeightTo").Value))
 


  SystemOptions.IsGeometricProportions = IIf(rs("IsGeometricProportions").Value = 0 Or IsNull(rs("IsGeometricProportions").Value), False, True)
  
  SystemOptions.CanPartialpayment = IIf(rs("CanPartialpayment").Value = 0 Or IsNull(rs("CanPartialpayment").Value), False, True)
  
  SystemOptions.EndRentifPayed = IIf(rs("EndRentifPayed").Value = 0 Or IsNull(rs("EndRentifPayed").Value), False, True)
  SystemOptions.cantCahngeAkarinExpenses = IIf(rs("cantCahngeAkarinExpenses").Value = 0 Or IsNull(rs("cantCahngeAkarinExpenses").Value), False, True)
  
  SystemOptions.EmployeeSalaryBYBranch = IIf(rs("EmployeeSalaryBYBranch").Value = 0 Or IsNull(rs("EmployeeSalaryBYBranch").Value), False, True)
  SystemOptions.returnnotcreatvoucher = IIf(rs("returnnotcreatvoucher").Value = 0 Or IsNull(rs("returnnotcreatvoucher").Value), False, True)
  
  SystemOptions.WaiverSetByContract = IIf(rs("WaiverSetByContract").Value = 0 Or IsNull(rs("WaiverSetByContract").Value), False, True)
SystemOptions.WaiverSetByContract = IIf(rs("WaiverSetByContract").Value = 0 Or IsNull(rs("WaiverSetByContract").Value), False, True)


  
   SystemOptions.CarsRevenuePerOwner = IIf(rs("CarsRevenuePerOwner").Value = 0 Or IsNull(rs("CarsRevenuePerOwner").Value), False, True)
  
  SystemOptions.IsCustSalesManCashRelated = IIf(rs("IsCustSalesManCashRelated").Value = 0 Or IsNull(rs("IsCustSalesManCashRelated").Value), False, True)
  SystemOptions.showEmployeeAccountIntrip = IIf(rs("showEmployeeAccountIntrip").Value = 0 Or IsNull(rs("showEmployeeAccountIntrip").Value), False, True)
  SystemOptions.DUEDOCUMENTbyinstallDate = IIf(rs("DUEDOCUMENTbyinstallDate").Value = 0 Or IsNull(rs("DUEDOCUMENTbyinstallDate").Value), False, True)
  
  SystemOptions.CanSkipPurchOrder = IIf(rs("CanSkipPurchOrder").Value = 0 Or IsNull(rs("CanSkipPurchOrder").Value), False, True)
  
  SystemOptions.CompilingBasedTable = IIf(rs("CompilingBasedTable").Value = 0 Or IsNull(rs("CompilingBasedTable").Value), False, True)
  SystemOptions.DontSaveInvoiceWithoutDocType = IIf(rs("DontSaveInvoiceWithoutDocType").Value = 0 Or IsNull(rs("DontSaveInvoiceWithoutDocType").Value), False, True)
  
  SystemOptions.DontDuplicateManulaNoInPurchase = IIf(rs("DontDuplicateManulaNoInPurchase").Value = 0 Or IsNull(rs("DontDuplicateManulaNoInPurchase").Value), False, True)
  
  
  SystemOptions.InvoiceTransferJLTotal = IIf(rs("InvoiceTransferJLTotal").Value = 0 Or IsNull(rs("InvoiceTransferJLTotal").Value), False, True)
  
      SystemOptions.EmpProduction = IIf(rs("EmpProduction").Value = 0 Or IsNull(rs("EmpProduction").Value), False, True)
    SystemOptions.ItemProduction = IIf(rs("ItemProduction").Value = 0 Or IsNull(rs("ItemProduction").Value), False, True)
    SystemOptions.ExpProduction = IIf(rs("ExpProduction").Value = 0 Or IsNull(rs("ExpProduction").Value), False, True)
    
SystemOptions.VATNoAccordActivity = IIf(rs("VATNoAccordActivity").Value = 0 Or IsNull(rs("VATNoAccordActivity").Value), False, True)
SystemOptions.NotCrtResvVouchProjects = IIf(rs("NotCrtResvVouchProjects").Value = 0 Or IsNull(rs("NotCrtResvVouchProjects").Value), False, True)


SystemOptions.LinkUsersWithPayment = IIf(rs("LinkUsersWithPayment").Value = 0 Or IsNull(rs("LinkUsersWithPayment").Value), False, True)

 SystemOptions.logowidth = IIf(rs("logowidth").Value = 0 Or IsNull(rs("logowidth").Value), 4000, rs("logowidth").Value)
 SystemOptions.logoHeight = IIf(rs("logoHeight").Value = 0 Or IsNull(rs("logoHeight").Value), 1500, rs("logoHeight").Value)
 
 
 
 
 SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").Value = 0 Or IsNull(rs("CustomerhavethreeAccounts").Value), False, True)
 
 SystemOptions.CustomerhavethreeAccounts = IIf(rs("CustomerhavethreeAccounts").Value = 0 Or IsNull(rs("CustomerhavethreeAccounts").Value), False, True)
 
 
    SystemOptions.CreateDriverBox = IIf(rs("CreateDriverBox").Value = 0 Or IsNull(rs("CreateDriverBox").Value), False, True)
    SystemOptions.CreateDriverEra = IIf(rs("CreateDriverEra").Value = 0 Or IsNull(rs("CreateDriverEra").Value), False, True)
 
    SystemOptions.TypicalProduction = IIf(rs("TypicalProduction").Value = 0 Or IsNull(rs("TypicalProduction").Value = 0), False, True)
 
    SystemOptions.ExpensesCoding = IIf(rs("ExpensesCoding").Value = 0 Or IsNull(rs("ExpensesCoding").Value), False, True)
    SystemOptions.ExpensesCoding2 = IIf(rs("ExpensesCoding2").Value = 0 Or IsNull(rs("ExpensesCoding2").Value), False, True)
    SystemOptions.SMSUserName = IIf(IsNull(rs("SMSUserName").Value), "", rs("SMSUserName").Value)
    SystemOptions.SMSPassWord = IIf(IsNull(rs("SMSPassWord").Value), "", rs("SMSPassWord").Value)
     SystemOptions.SenderName = IIf(IsNull(rs("SenderName").Value), "", rs("SenderName").Value)
   SystemOptions.OPTWEB = IIf(IsNull(rs("optweb").Value), 0, rs("optweb").Value)
   
    SystemOptions.InstallmntsvchrCoding = IIf(rs("InstallmntsvchrCoding").Value = 0 Or IsNull(rs("InstallmntsvchrCoding").Value), False, True)
 
    SystemOptions.AllowIndirectCost = IIf(rs("AllowIndirectCost").Value = 0 Or IsNull(rs("AllowIndirectCost").Value), False, True)
 
    SystemOptions.banks_Accounts3 = IIf(rs("banks_Accounts").Value = 0 Or IsNull(rs("banks_Accounts").Value), False, True)
    SystemOptions.AssetAccount = IIf(rs("AssetAccount").Value = 0 Or IsNull(rs("AssetAccount").Value), False, True)
    SystemOptions.AssetAccount1 = IIf(rs("AssetAccount1").Value = 0 Or IsNull(rs("AssetAccount1").Value), False, True)
 '**********************************************************************
     SystemOptions.StoreAccountHaveSettelment = IIf(rs("StoreAccountHaveSettelment").Value = 0 Or IsNull(rs("StoreAccountHaveSettelment").Value), False, True)
     SystemOptions.eachStoreHaveLossAccount = IIf(IsNull(rs("eachStoreHaveLossAccount").Value), True, IIf((rs("eachStoreHaveLossAccount").Value = 0), False, True))
     SystemOptions.eachStoreHaveGiftAccount = IIf(IsNull(rs("eachStoreHaveGiftAccount").Value), True, IIf((rs("eachStoreHaveGiftAccount").Value = 0), False, True))
             
 '**********************************************************************
    SystemOptions.autoReseiveVoucher = IIf(rs("autoReseiveVoucher").Value = 0 Or IsNull(rs("autoReseiveVoucher").Value), False, True)
  
   SystemOptions.ReturnSallingOption = IIf(IsNull(rs("ReturnSallingOption").Value), False, (rs("ReturnSallingOption").Value))
    SystemOptions.ReturnSallingIntervalCount = IIf(IsNull(rs("ReturnSallingIntervalCount").Value), 0, Val(rs("ReturnSallingIntervalCount").Value))
    SystemOptions.ReturnSallingIntervalCount1 = IIf(IsNull(rs("ReturnSallingIntervalCount1").Value), 0, Val(rs("ReturnSallingIntervalCount1").Value))

    SystemOptions.DateOpt = IIf(IsNull(rs("DateOpt").Value), 0, Val(rs("DateOpt").Value))

    SystemOptions.IndirectCostPercentage = IIf(IsNull(rs("IndirectCostPercentage").Value), 0, Val(rs("IndirectCostPercentage").Value))
    
    SystemOptions.StoreDigit = IIf(IsNull(rs("StoreDigit").Value), 1, (rs("StoreDigit").Value))
    SystemOptions.BranchDigit = IIf(IsNull(rs("BranchDigit").Value), 1, (rs("BranchDigit").Value))
    
    
    SystemOptions.Ked_digit = IIf(IsNull(rs("Ked_digit").Value), 0, Val(rs("Ked_digit").Value))
    SystemOptions.Count_ACCOUNT_digit = IIf(IsNull(rs("Count_ACCOUNT_digit").Value), 2, Val(rs("Count_ACCOUNT_digit").Value))
    SystemOptions.Save_options = IIf(IsNull(rs("Save_options").Value), 0, Val(rs("Save_options").Value))
    SystemOptions.ReservEmp = IIf(IsNull(rs("EmpRes").Value), 0, (rs("EmpRes").Value))

    SystemOptions.EmpComponentDigts = IIf(IsNull(rs("EmpComponentDigts").Value), 2, Val(rs("EmpComponentDigts").Value))
            SystemOptions.ImagesPath = IIf(IsNull(rs("ImagesPath").Value), "Images", (rs("ImagesPath").Value))
     SystemOptions.Reportpath = IIf(IsNull(rs("reportPath").Value), "Stander", (rs("reportPath").Value))
     SystemOptions.BigUserPw = IIf(IsNull(rs("BigUserPw").Value), "n20172018", (rs("BigUserPw").Value))
   If SystemOptions.BigUserPw = "" Then SystemOptions.BigUserPw = "n20172018"
     'BigUserPw
     

SystemOptions.CostStarting = IIf(rs("CostStarting").Value = 0 Or IsNull(rs("CostStarting").Value), False, True)

SystemOptions.chkuserCode = IIf(rs("chkuserCode").Value = 0 Or IsNull(rs("chkuserCode").Value), False, True)
SystemOptions.Itemsattachedzero = IIf(rs("Itemsattachedzero").Value = 0 Or IsNull(rs("Itemsattachedzero").Value), False, True)
SystemOptions.workWithBarcode = IIf(rs("workWithBarcode").Value = 0 Or IsNull(rs("workWithBarcode").Value), False, True)
SystemOptions.WorkWithBarCodeParent = IIf(rs("WorkWithBarCodeParent").Value = 0 Or IsNull(rs("WorkWithBarCodeParent").Value), False, True)

SystemOptions.WorkWithLINKEDiTEMS = IIf(rs("WorkWithLINKEDiTEMS").Value = 0 Or IsNull(rs("WorkWithLINKEDiTEMS").Value), False, True)
SystemOptions.WorkWithBranchLogo = IIf(rs("WorkWithBranchLogo").Value = 0 Or IsNull(rs("WorkWithBranchLogo").Value), False, True)

SystemOptions.WorkWithFirstInstallOnly = IIf(rs("WorkWithFirstInstallOnly").Value = 0 Or IsNull(rs("WorkWithFirstInstallOnly").Value), False, True)
SystemOptions.WorkWithGroupCode = IIf(rs("WorkWithGroupCode").Value = 0 Or IsNull(rs("WorkWithGroupCode").Value), False, True)
'31032017egypt

SystemOptions.AllowSalesMultyPayed = IIf(rs("AllowSalesMultyPayed").Value = 0 Or IsNull(rs("AllowSalesMultyPayed").Value), False, True)
SystemOptions.MultyStore = IIf(rs("MultyStore").Value = 0 Or IsNull(rs("MultyStore").Value), False, True)
SystemOptions.RawMaterMix2 = IIf(rs("RawMaterMix2").Value = 0 Or IsNull(rs("RawMaterMix2").Value), False, True)
 SystemOptions.DontShowMoreDetailsCompItem = IIf(rs("DontShowMoreDetailsCompItem").Value = 0 Or IsNull(rs("DontShowMoreDetailsCompItem").Value), False, True)
 SystemOptions.traveDiscountFromCustomerDirect = IIf(rs("traveDiscountFromCustomerDirect").Value = 0 Or IsNull(rs("traveDiscountFromCustomerDirect").Value), False, True)
 
 
SystemOptions.TransferNotInvItemDef = IIf(rs("TransferNotInvItemDef").Value = 0 Or IsNull(rs("TransferNotInvItemDef").Value), False, True)


SystemOptions.CustMobNoMandatory = IIf(rs("CustMobNoMandatory").Value = 0 Or IsNull(rs("CustMobNoMandatory").Value), False, True)

    SystemOptions.CostProductOrderByOut = IIf(rs("CostProductOrderByOut").Value = 0 Or IsNull(rs("CostProductOrderByOut").Value), False, True)
SystemOptions.SortInvoiceByEntry = IIf(rs("SortInvoiceByEntry").Value = 0 Or IsNull(rs("SortInvoiceByEntry").Value), False, True)






    


SystemOptions.CashCustomerNameMustenter = IIf(rs("CashCustomerNameMustenter").Value = 0 Or IsNull(rs("CashCustomerNameMustenter").Value), False, True)
SystemOptions.AllowCommtionJEFromValueVisa = IIf(rs("AllowCommtionJEFromValueVisa").Value = 0 Or IsNull(rs("AllowCommtionJEFromValueVisa").Value), False, True)
SystemOptions.AllowWorkWithArea = IIf(rs("AllowWorkWithArea").Value = 0 Or IsNull(rs("AllowCommtionJEFromValueVisa").Value), False, True)
SystemOptions.AllowPurchasesMultyPayed = IIf(rs("AllowPurchasesMultyPayed").Value = 0 Or IsNull(rs("AllowPurchasesMultyPayed").Value), False, True)
SystemOptions.AllowDynamicEdit = False
SystemOptions.AllowDynamicEdit = IIf(rs("AllowDynamicEdit").Value = 0 Or IsNull(rs("AllowDynamicEdit").Value), False, True)
SystemOptions.AllowDynamicAutoSus = False
SystemOptions.AllowDynamicAutoSus = IIf(rs("AllowDynamicAutoSus").Value = 0 Or IsNull(rs("AllowDynamicAutoSus").Value), False, True)

SystemOptions.AllowUnbalncedByBranchAccount = False
SystemOptions.AllowUnbalncedByBranchAccount = IIf(rs("AllowUnbalncedByBranchAccount").Value = 0 Or IsNull(rs("AllowUnbalncedByBranchAccount").Value), False, True)



SystemOptions.CloseMovingVchrinSales = False
SystemOptions.CloseMovingVchrinSales = IIf(rs("CloseMovingVchrinSales").Value = 0 Or IsNull(rs("CloseMovingVchrinSales").Value), False, True)


SystemOptions.CantChangeSalesPerson = False
SystemOptions.CantChangeSalesPerson = IIf(rs("CantChangeSalesPerson").Value = 0 Or IsNull(rs("CantChangeSalesPerson").Value), False, True)

SystemOptions.BatchCreateManyworkOrder = False
SystemOptions.BatchCreateManyworkOrder = IIf(rs("BatchCreateManyworkOrder").Value = 0 Or IsNull(rs("BatchCreateManyworkOrder").Value), False, True)




SystemOptions.AllItemInVAT = False
SystemOptions.AllItemInVAT = IIf(rs("AllItemInVAT").Value = 0 Or IsNull(rs("AllItemInVAT").Value), False, True)
SystemOptions.SendToAprovedSalesBill = IIf(rs("SendToAprovedSalesBill").Value = 0 Or IsNull(rs("SendToAprovedSalesBill").Value), False, True)
SystemOptions.SalaryJLByAnalyEqup = IIf(rs("SalaryJLByAnalyEqup").Value = 0 Or IsNull(rs("SalaryJLByAnalyEqup").Value), False, True)


SystemOptions.AllowSaveTripWithoutExpen = IIf(rs("AllowSaveTripWithoutExpen").Value = 0 Or IsNull(rs("AllowSaveTripWithoutExpen").Value), False, True)
SystemOptions.SAVEMAINTENANCEJOBWITHORDERORPLANONLY = IIf(rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").Value = 0 Or IsNull(rs("SAVEMAINTENANCEJOBWITHORDERORPLANONLY").Value), False, True)

 SystemOptions.TransBillPriceByGrid = IIf(rs("TransBillPriceByGrid").Value = 0 Or IsNull(rs("TransBillPriceByGrid").Value), False, True)
SystemOptions.NoCreatJLInRentContract = IIf(rs("NoCreatJLInRentContract").Value = 0 Or IsNull(rs("NoCreatJLInRentContract").Value), False, True)
SystemOptions.OpenVATAccountOwner = IIf(rs("OpenVATAccountOwner").Value = 0 Or IsNull(rs("OpenVATAccountOwner").Value), False, True)


SystemOptions.CreateJLEmpCommissions = IIf(rs("CreateJLEmpCommissions").Value = 0 Or IsNull(rs("CreateJLEmpCommissions").Value), False, True)
SystemOptions.TypeContractAutoFromIqar = IIf(rs("TypeContractAutoFromIqar").Value = 0 Or IsNull(rs("TypeContractAutoFromIqar").Value), False, True)
SystemOptions.AllowRepeatInvoiceNo = IIf(rs("AllowRepeatInvoiceNo").Value = 0 Or IsNull(rs("AllowRepeatInvoiceNo").Value), False, True)
SystemOptions.EmpSalaryDigts = IIf(IsNull(rs("EmpSalaryDigts").Value), 2, Val(rs("EmpSalaryDigts").Value))
SystemOptions.AllowReturnFIFO = IIf(rs("AllowReturnFIFO").Value = 0 Or IsNull(rs("AllowReturnFIFO").Value), False, True)
SystemOptions.AllowDiscountAllowedFIFO = IIf(rs("AllowDiscountAllowedFIFO").Value = 0 Or IsNull(rs("AllowDiscountAllowedFIFO").Value), False, True)
SystemOptions.AllowJLManualFIFO = IIf(rs("AllowJLManualFIFO").Value = 0 Or IsNull(rs("AllowJLManualFIFO").Value), False, True)




SystemOptions.EmpSalaryDigts = IIf(IsNull(rs("EmpSalaryDigts").Value), 2, Val(rs("EmpSalaryDigts").Value))
SystemOptions.AllowReturnFIFO = IIf(rs("AllowReturnFIFO").Value = 0 Or IsNull(rs("AllowReturnFIFO").Value), False, True)
SystemOptions.AllowDiscountAllowedFIFO = IIf(rs("AllowDiscountAllowedFIFO").Value = 0 Or IsNull(rs("AllowDiscountAllowedFIFO").Value), False, True)
SystemOptions.AllowJLManualFIFO = IIf(rs("AllowJLManualFIFO").Value = 0 Or IsNull(rs("AllowJLManualFIFO").Value), False, True)



SystemOptions.PaymentIntoAccouStat = IIf(rs("PaymentIntoAccouStat").Value = 0 Or IsNull(rs("PaymentIntoAccouStat").Value), False, True)
SystemOptions.ProvisionsByManagement = IIf(rs("ProvisionsByManagement").Value = 0 Or IsNull(rs("ProvisionsByManagement").Value), False, True)

SystemOptions.ProvisionsByıEQuipments = IIf(rs("ProvisionsByıEQuipments").Value = 0 Or IsNull(rs("ProvisionsByıEQuipments").Value), False, True)

SystemOptions.ReturnSAlesByBarcode = IIf(rs("ReturnSAlesByBarcode").Value = 0 Or IsNull(rs("ReturnSAlesByBarcode").Value), False, True)
SystemOptions.CreatePayOrderSales = IIf(rs("CreatePayOrderSales").Value = 0 Or IsNull(rs("CreatePayOrderSales").Value), False, True)
SystemOptions.TripnotUploadExpenses = IIf(rs("TripnotUploadExpenses").Value = 0 Or IsNull(rs("TripnotUploadExpenses").Value), False, True)
SystemOptions.ExpensesByQtyOnly = IIf(rs("ExpensesByQtyOnly").Value = 0 Or IsNull(rs("ExpensesByQtyOnly").Value), False, True)

SystemOptions.ShowPrinterDialoge = IIf(rs("ShowPrinterDialoge").Value = 0 Or IsNull(rs("ShowPrinterDialoge").Value), False, True)





SystemOptions.IsBarCodeByUnit = IIf(rs("IsBarCodeByUnit").Value = 0 Or IsNull(rs("IsBarCodeByUnit").Value), False, True)


SystemOptions.DontDistributeLegalACC = IIf(rs("DontDistributeLegalACC").Value = 0 Or IsNull(rs("DontDistributeLegalACC").Value), False, True)

 

SystemOptions.AllowEditInvoiceNoticeDiscount = IIf(rs("AllowEditInvoiceNoticeDiscount").Value = 0 Or IsNull(rs("AllowEditInvoiceNoticeDiscount").Value), False, True)
SystemOptions.AllowEditInvoiceOfReturn = IIf(rs("AllowEditInvoiceOfReturn").Value = 0 Or IsNull(rs("AllowEditInvoiceOfReturn").Value), False, True)

SystemOptions.IsMultiItemsInCompItem = False
SystemOptions.IsMultiItemsInCompItem = IIf(rs("IsMultiItemsInCompItem").Value = 0 Or IsNull(rs("IsMultiItemsInCompItem").Value), False, True)

SystemOptions.LimitDefaultCredit = IIf(IsNull(rs("LimitDefaultCredit").Value), 0, Val(rs("LimitDefaultCredit").Value))
SystemOptions.LimitDefaultCreditDays = IIf(IsNull(rs("LimitDefaultCreditDays").Value), 0, Val(rs("LimitDefaultCreditDays").Value))




SystemOptions.ShowBalanceOfEmpInSalary = IIf(rs("ShowBalanceOfEmpInSalary").Value = 0 Or IsNull(rs("ShowBalanceOfEmpInSalary").Value), False, True)
SystemOptions.AllowScInterface = False

 SystemOptions.DontCreateOut = IIf(rs("DontCreateOut").Value = 0 Or IsNull(rs("DontCreateOut").Value), False, True)
 SystemOptions.DontCreateOut2 = IIf(rs("DontCreateOut2").Value = 0 Or IsNull(rs("DontCreateOut2").Value), False, True)
 SystemOptions.InsertItemManualOut = IIf(rs("InsertItemManualOut").Value = 0 Or IsNull(rs("InsertItemManualOut").Value), False, True)
    


SystemOptions.OpenAccountAqar = IIf(rs("OpenAccountAqar").Value = 0 Or IsNull(rs("OpenAccountAqar").Value), False, True)

SystemOptions.ShowOnlyItemsOfSales = False
SystemOptions.ShowOnlyItemsOfSales = IIf(rs("ShowOnlyItemsOfSales").Value = 0 Or IsNull(rs("ShowOnlyItemsOfSales").Value), False, True)

SystemOptions.PrintInvoiceByBranch = False
SystemOptions.PrintInvoiceByBranch = IIf(rs("PrintInvoiceByBranch").Value = 0 Or IsNull(rs("PrintInvoiceByBranch").Value), False, True)

SystemOptions.GeneralVoucherCreateSalesGE = False
SystemOptions.GeneralVoucherCreateSalesGE = IIf(rs("GeneralVoucherCreateSalesGE").Value = 0 Or IsNull(rs("GeneralVoucherCreateSalesGE").Value), False, True)


SystemOptions.SalesNotCreateGe = False
SystemOptions.SalesNotCreateGe = IIf(rs("SalesNotCreateGe").Value = 0 Or IsNull(rs("SalesNotCreateGe").Value), False, True)





SystemOptions.LinkSupplerWithItem = False
SystemOptions.LinkSupplerWithItem = IIf(rs("LinkSupplerWithItem").Value = 0 Or IsNull(rs("LinkSupplerWithItem").Value), False, True)

SystemOptions.NotAllowedCalcVata = IIf(rs("NotAllowedCalcVata").Value = 0 Or IsNull(rs("NotAllowedCalcVata").Value), False, True)
SystemOptions.LinkCustomerWithCars = IIf(rs("LinkCustomerWithCars").Value = 0 Or IsNull(rs("LinkCustomerWithCars").Value), False, True)

SystemOptions.AllowEditCashingLinkProj = IIf(rs("AllowEditCashingLinkProj").Value = 0 Or IsNull(rs("AllowEditCashingLinkProj").Value), False, True)



SystemOptions.AllowScInterface = IIf(rs("AllowScInterface").Value = 0 Or IsNull(rs("AllowScInterface").Value), False, True)

SystemOptions.IssueVoucherWorkWithRemain = IIf(rs("IssueVoucherWorkWithRemain").Value = 0 Or IsNull(rs("IssueVoucherWorkWithRemain").Value), False, True)
SystemOptions.TripDateInsertDefulat = IIf(rs("TripDateInsertDefulat").Value = 0 Or IsNull(rs("TripDateInsertDefulat").Value), False, True)
SystemOptions.TripwithorderOnly = IIf(rs("TripwithorderOnly").Value = 0 Or IsNull(rs("TripwithorderOnly").Value), False, True)

SystemOptions.AllowPriceWithWidth = IIf(rs("AllowPriceWithWidth").Value = 0 Or IsNull(rs("AllowPriceWithWidth").Value), False, True)



SystemOptions.DealingWithPrepayAccount = IIf(rs("DealingWithPrepayAccount").Value = 0 Or IsNull(rs("DealingWithPrepayAccount").Value), False, True)
SystemOptions.CreateJLVactionAratha = IIf(rs("CreateJLVactionAratha").Value = 0 Or IsNull(rs("CreateJLVactionAratha").Value), False, True)
SystemOptions.PriceWithVAT = IIf(rs("PriceWithVAT").Value = 0 Or IsNull(rs("PriceWithVAT").Value), False, True)
   SystemOptions.AllowWorkCustomerPoints = IIf(rs("AllowWorkCustomerPoints").Value = 0 Or IsNull(rs("AllowWorkCustomerPoints").Value), False, True)
SystemOptions.ProjectInvoiceAnalysisJL = IIf(rs("ProjectInvoiceAnalysisJL").Value = 0 Or IsNull(rs("ProjectInvoiceAnalysisJL").Value), False, True)

SystemOptions.CustomerRecordNoIsnotManda = IIf(rs("CustomerRecordNoIsnotManda").Value = 0 Or IsNull(rs("CustomerRecordNoIsnotManda").Value), False, True)

SystemOptions.DueComm = IIf(rs("DueComm").Value = 0 Or IsNull(rs("DueComm").Value), False, True)
SystemOptions.DueWater = IIf(rs("DueWater").Value = 0 Or IsNull(rs("DueWater").Value), False, True)
SystemOptions.DueElectr = IIf(rs("DueElectr").Value = 0 Or IsNull(rs("DueElectr").Value), False, True)
SystemOptions.DueService = IIf(rs("DueService").Value = 0 Or IsNull(rs("DueService").Value), False, True)
SystemOptions.CommissionOnOwner = IIf(rs("CommissionOnOwner").Value = 0 Or IsNull(rs("CommissionOnOwner").Value), False, True)

SystemOptions.CommissionDue = IIf(rs("CommissionDue").Value = 0 Or IsNull(rs("CommissionDue").Value), False, True)
SystemOptions.SupplierReciveGE = IIf(rs("SupplierReciveGE").Value = 0 Or IsNull(rs("SupplierReciveGE").Value), False, True)
SystemOptions.BranchmustimSalary = IIf(rs("BranchmustimSalary").Value = 0 Or IsNull(rs("BranchmustimSalary").Value), False, True)



SystemOptions.InsuranceOnOwner = IIf(rs("InsuranceOnOwner").Value = 0 Or IsNull(rs("InsuranceOnOwner").Value), False, True)
SystemOptions.ServicesOnOwner = IIf(rs("ServicesOnOwner").Value = 0 Or IsNull(rs("ServicesOnOwner").Value), False, True)
SystemOptions.AllowProductOrderOne = IIf(rs("AllowProductOrderOne").Value = 0 Or IsNull(rs("AllowProductOrderOne").Value), False, True)
SystemOptions.SalaryJLByManagement = IIf(rs("SalaryJLByManagement").Value = 0 Or IsNull(rs("SalaryJLByManagement").Value), False, True)



SystemOptions.AllowChangePriceApprove = IIf(rs("AllowChangePriceApprove").Value = 0 Or IsNull(rs("AllowChangePriceApprove").Value), False, True)
SystemOptions.AllowSkipPayment = IIf(rs("AllowSkipPayment").Value = 0 Or IsNull(rs("AllowSkipPayment").Value), False, True)


SystemOptions.AllowAnalyticJL = IIf(rs("AllowAnalyticJL").Value = 0 Or IsNull(rs("AllowAnalyticJL").Value), False, True)

SystemOptions.AllowGoodPerfAccount = IIf(rs("AllowGoodPerfAccount").Value = 0 Or IsNull(rs("AllowGoodPerfAccount").Value), False, True)
SystemOptions.ManualSalesInvoiceMust = IIf(rs("ManualSalesInvoiceMust").Value = 0 Or IsNull(rs("ManualSalesInvoiceMust").Value), False, True)


 
'
SystemOptions.SalesTrustsAffectVedorCode = IIf(rs("SalesTrustsAffectVedorCode").Value = 0 Or IsNull(rs("SalesTrustsAffectVedorCode").Value), False, True)
  
  
  
SystemOptions.AllowItemByRow = IIf(rs("AllowItemByRow").Value = 0 Or IsNull(rs("AllowItemByRow").Value), False, True)
SystemOptions.AllowChangManualQtyMix = IIf(rs("AllowChangManualQtyMix").Value = 0 Or IsNull(rs("AllowChangManualQtyMix").Value), False, True)
SystemOptions.AccountAccordingCash = IIf(rs("AccountAccordingCash").Value = 0 Or IsNull(rs("AccountAccordingCash").Value), False, True)



SystemOptions.ProductionRawMaterMix = IIf(rs("ProductionRawMaterMix").Value = 0 Or IsNull(rs("ProductionRawMaterMix").Value), False, True)
SystemOptions.AllowLastPrice = IIf(rs("AllowLastPrice").Value = 0 Or IsNull(rs("AllowLastPrice").Value), False, True)

SystemOptions.AllowAcceleratepayment = IIf(rs("AllowAcceleratepayment").Value = 0 Or IsNull(rs("AllowAcceleratepayment").Value), False, True)
SystemOptions.AllowExperDateFIFO = IIf(rs("AllowExperDateFIFO").Value = 0 Or IsNull(rs("AllowExperDateFIFO").Value), False, True)
SystemOptions.AllowProjectBill2Serial = IIf(rs("AllowProjectBill2Serial").Value = 0 Or IsNull(rs("AllowProjectBill2Serial").Value), False, True)
SystemOptions.AllowNoRoudProjectInvoices = IIf(rs("AllowNoRoudProjectInvoices").Value = 0 Or IsNull(rs("AllowNoRoudProjectInvoices").Value), False, True)




SystemOptions.NOOFPRINTCOPIESSALES = IIf(rs("NOOFPRINTCOPIESSALES").Value = 0 Or IsNull(rs("NOOFPRINTCOPIESSALES").Value), 0, rs("NOOFPRINTCOPIESSALES").Value)


'31032017egypt

SystemOptions.DecideItemName = IIf(rs("DecideItemName").Value = 0 Or IsNull(rs("DecideItemName").Value), False, True)
SystemOptions.DefaultIsCreditSales = IIf(rs("DefaultIsCreditSales").Value = 0 Or IsNull(rs("DefaultIsCreditSales").Value), False, True)
SystemOptions.DefaultIsCreditPurchase = IIf(rs("DefaultIsCreditPurchase").Value = 0 Or IsNull(rs("DefaultIsCreditSales").Value), False, True)

SystemOptions.DefaultIsCreditPurchaseRet = IIf(rs("DefaultIsCreditPurchaseRet").Value = 0 Or IsNull(rs("DefaultIsCreditPurchaseRet").Value), False, True)

SystemOptions.returnByBarCodeOnly = IIf(rs("returnByBarCodeOnly").Value = 0 Or IsNull(rs("returnByBarCodeOnly").Value), False, True)


SystemOptions.JLCodeBasedOnBranch = IIf(rs("JLCodeBasedOnBranch").Value = 0 Or IsNull(rs("JLCodeBasedOnBranch").Value), False, True)

SystemOptions.EmpNotExcceedDiscount = IIf(rs("EmpNotExcceedDiscount").Value = 0 Or IsNull(rs("EmpNotExcceedDiscount").Value), False, True)

SystemOptions.BoxLossandIncreae = IIf(rs("BoxLossandIncreae").Value = 0 Or IsNull(rs("BoxLossandIncreae").Value), False, True)


SystemOptions.attacheditemsisfree = IIf(rs("attacheditemsisfree").Value = 0 Or IsNull(rs("attacheditemsisfree").Value), False, True)
 


If IsNull(rs("EnableCustomerAging").Value) Then
SystemOptions.EnableCustomerAging = True
Else
SystemOptions.EnableCustomerAging = IIf(rs("EnableCustomerAging").Value = 0, False, True)
End If


If IsNull(rs("showcostColorininvoice").Value) Then
SystemOptions.showcostColorininvoice = True
Else
SystemOptions.showcostColorininvoice = IIf(rs("showcostColorininvoice").Value = 0, False, True)
End If

If IsNull(rs("SubContactorHave3Account").Value) Then
SystemOptions.SubContactorHave3Account = False
Else
SystemOptions.SubContactorHave3Account = IIf(rs("SubContactorHave3Account").Value = 0, False, True)
End If

If IsNull(rs("ProjectEmployeeGV").Value) Then
SystemOptions.ProjectEmployeeGV = False
Else
SystemOptions.ProjectEmployeeGV = IIf(rs("ProjectEmployeeGV").Value = 0, False, True)
End If

If IsNull(rs("PursgaseWithoutDecimal").Value) Then
SystemOptions.PursgaseWithoutDecimal = False
Else
SystemOptions.PursgaseWithoutDecimal = IIf(rs("PursgaseWithoutDecimal").Value = 0, False, True)
End If


If IsNull(rs("workWithCustomerContract").Value) Then
SystemOptions.workWithCustomerContract = False
Else
SystemOptions.workWithCustomerContract = IIf(rs("workWithCustomerContract").Value = 0, False, True)
End If


If IsNull(rs("workWithVendorContract").Value) Then
SystemOptions.workWithvendorContract = False
Else
SystemOptions.workWithvendorContract = IIf(rs("workWithVendorContract").Value = 0, False, True)
End If


If IsNull(rs("PoCreateVoucher").Value) Then
SystemOptions.PoCreateVoucher = False
Else
SystemOptions.PoCreateVoucher = IIf(rs("PoCreateVoucher").Value = 0, False, True)
End If


If IsNull(rs("DiscountSalesCreateVchr").Value) Then
SystemOptions.DiscountSalesCreateVchr = False
Else
SystemOptions.DiscountSalesCreateVchr = IIf(rs("DiscountSalesCreateVchr").Value = 0, False, True)
End If

SystemOptions.AllowCostnNewShape = False
 
 
 If IsNull(rs("AllowCostnNewShape").Value) Then
    SystemOptions.AllowCostnNewShape = False
Else
    SystemOptions.AllowCostnNewShape = IIf(rs("AllowCostnNewShape").Value = 0, False, True)
End If


SystemOptions.ProjectUnderImplemen = False
SystemOptions.ProjectUnderImplemen = IIf(rs("ProjectUnderImplemen").Value = 0 Or IsNull(rs("ProjectUnderImplemen").Value), False, True)


SystemOptions.AllowCostBySerial = False
 
 
 If IsNull(rs("AllowCostBySerial").Value) Then
    SystemOptions.AllowCostBySerial = False
Else
    SystemOptions.AllowCostBySerial = IIf(rs("AllowCostBySerial").Value = 0, False, True)
End If


If IsNull(rs("AllowCostPerStore").Value) Then
SystemOptions.AllowCostPerStore = False
Else
SystemOptions.AllowCostPerStore = IIf(rs("AllowCostPerStore").Value = 0, False, True)
End If

''DB_CreateField "TblOptions", "AllowCostPerStore", adBoolean, adColNullable, , , "                ", False, True
If IsNull(rs("PaymentDifferent").Value) Then
SystemOptions.PaymentDifferent = False
Else
SystemOptions.PaymentDifferent = IIf(rs("PaymentDifferent").Value = 0, False, True)
End If


 

If IsNull(rs("poWithatotalQty").Value) Then
SystemOptions.poWithatotalQty = False
Else
SystemOptions.poWithatotalQty = IIf(rs("poWithatotalQty").Value = 0, False, True)
End If


If IsNull(rs("PayrollOneAccount").Value) Then
SystemOptions.PayrollOneAccount = False
Else
SystemOptions.PayrollOneAccount = IIf(rs("PayrollOneAccount").Value = 0, False, True)
End If

If IsNull(rs("WorkWithItemsDetails").Value) Then
SystemOptions.WorkWithItemsDetails = False
Else
SystemOptions.WorkWithItemsDetails = IIf(rs("WorkWithItemsDetails").Value = 0, False, True)
End If


If IsNull(rs("FAAddtionCreateAccount").Value) Then
SystemOptions.FAAddtionCreateAccount = False
Else
SystemOptions.FAAddtionCreateAccount = IIf(rs("FAAddtionCreateAccount").Value = 0, False, True)
End If

 If IsNull(rs("Create2account4Supp").Value) Then
SystemOptions.Create2account4Supp = False
Else
SystemOptions.Create2account4Supp = IIf(rs("Create2account4Supp").Value = 0, False, True)
End If



'DB_CreateField "TblOptions", "poWithatotalQty", adBoolean, adColNullable, , , "                ", False, True


If IsNull(rs("cancellAllApprove").Value) Then
SystemOptions.cancellAllApprove = False
Else
SystemOptions.cancellAllApprove = IIf(rs("cancellAllApprove").Value = 0, False, True)
End If


If IsNull(rs("workwithticketAllocation").Value) Then
SystemOptions.workwithticketAllocation = False
Else
SystemOptions.workwithticketAllocation = IIf(rs("workwithticketAllocation").Value = 0, False, True)
End If

'PursgaseWithoutDecimal


'''DB_CreateField "TblOptions", "JLCodeBasedOnBranch", adBoolean, adColNullable, , , " «·«› —«÷Ì «·»Ì⁄ «Ã·          ", False, True

SystemOptions.AnalyticPaymentVouchr = IIf(rs("AnalyticPaymentVouchr").Value = 0 Or IsNull(rs("AnalyticPaymentVouchr").Value), False, True)

SystemOptions.ShowDriverOnly = IIf(rs("ShowDriverOnly").Value = 0 Or IsNull(rs("ShowDriverOnly").Value), False, True)

SystemOptions.CreateInsuranceAccountForCustomers = IIf(rs("CreateInsuranceAccountForCustomers").Value = 0 Or IsNull(rs("CreateInsuranceAccountForCustomers").Value), False, True)

SystemOptions.DuplicateitemsNames = IIf(rs("DuplicateitemsNames").Value = 0 Or IsNull(rs("DuplicateitemsNames").Value), False, True)

SystemOptions.TradingPOS = IIf(rs("TradingPOS").Value = 0 Or IsNull(rs("TradingPOS").Value), False, True)
SystemOptions.posshape2 = IIf(rs("posshape2").Value = 0 Or IsNull(rs("posshape2").Value), False, True)

SystemOptions.CanChanegeLinkedSsalesnvoice = IIf(rs("CanChanegeLinkedSsalesnvoice").Value = 0 Or IsNull(rs("CanChanegeLinkedSsalesnvoice").Value), False, True)

SystemOptions.CanChanegeLinkedPurcahsenvoice = IIf(rs("CanChanegeLinkedPurcahsenvoice").Value = 0 Or IsNull(rs("CanChanegeLinkedPurcahsenvoice").Value), False, True)

 SystemOptions.updatecashvchrifdeposite = IIf(rs("updatecashvchrifdeposite").Value = 0 Or IsNull(rs("updatecashvchrifdeposite").Value), False, True)

SystemOptions.Revenueowed = IIf(rs("Revenueowed").Value = 0 Or IsNull(rs("Revenueowed").Value), False, True)
SystemOptions.AllowupdateJobStatus = IIf(rs("AllowupdateJobStatus").Value = 0 Or IsNull(rs("AllowupdateJobStatus").Value), False, True)

SystemOptions.OpeningEmployeeShowAll = IIf(rs("OpeningEmployeeShowAll").Value = 0 Or IsNull(rs("OpeningEmployeeShowAll").Value), False, True)
SystemOptions.EndServiceMore5Year = IIf(rs("EndServiceMore5Year").Value = 0 Or IsNull(rs("EndServiceMore5Year").Value), False, True)

SystemOptions.VacstionShowOldSalaries = IIf(rs("VacstionShowOldSalaries").Value = 0 Or IsNull(rs("VacstionShowOldSalaries").Value), False, True)
SystemOptions.AllowReturnWithoutCost = IIf(rs("AllowReturnWithoutCost").Value = 0 Or IsNull(rs("AllowReturnWithoutCost").Value), False, True)



'

'SALIMSystemOptions
SystemOptions.ShowItemByCustomer = IIf(rs("ShowItemByCustomer").Value = 0 Or IsNull(rs("ShowItemByCustomer").Value), False, True)
SystemOptions.AllowTowShift = IIf(rs("AllowTowShift").Value = 0 Or IsNull(rs("AllowTowShift").Value), False, True)

SystemOptions.AllowItemsShortName = IIf(rs("AllowItemsShortName").Value = 0 Or IsNull(rs("AllowItemsShortName").Value), False, True)
SystemOptions.SellOrderBalance = IIf(rs("SellOrderBalance").Value = 0 Or IsNull(rs("SellOrderBalance").Value), False, True)



 

SystemOptions.Ecnomy = IIf(rs("Ecnomy").Value = 0 Or IsNull(rs("Ecnomy").Value), False, True)
 SystemOptions.WebAdv = IIf(rs("WebAdv").Value = "" Or IsNull(rs("WebAdv").Value), "", rs("WebAdv").Value)
SystemOptions.ViewAccountsbyBranch = IIf(rs("ViewAccountsbyBranch").Value = 0 Or IsNull(rs("ViewAccountsbyBranch").Value), False, True)
SystemOptions.AllowEditeAccounts = IIf(rs("AllowEditeAccounts").Value = 0 Or IsNull(rs("AllowEditeAccounts").Value), False, True)
SystemOptions.AllowHideAssest = IIf(rs("AllowHideAssest").Value = 0 Or IsNull(rs("AllowHideAssest").Value), False, True)

SystemOptions.AllowAccountMultyPayed = IIf(rs("AllowAccountMultyPayed").Value = 0 Or IsNull(rs("AllowAccountMultyPayed").Value), False, True)


  SystemOptions.LockSalary = IIf(rs("LockSalary").Value = 0 Or IsNull(rs("LockSalary").Value), False, True)
    IntTemp = rs("CurrencyDigts").Value
    Decimal_Places = IntTemp

    IntTemp1 = rs("PriceDigtsInst").Value
    Decimal_Places1 = IntTemp1

    StrTemp = String$(IntTemp, "0")

    If StrTemp = "" Then
        SystemOptions.SysDefCurrencyForamt = IntTemp ' "" '"#,###"
    Else
        SystemOptions.SysDefCurrencyForamt = IntTemp ' "" '  "#,###." & StrTemp
    End If

    IntTemp = rs("QtyDigts").Value

    SystemOptions.SysDefQuantityDecimal = rs("QtyDigts").Value
    StrTemp = String$(IntTemp, "0")
    SystemOptions.SysDefQuantityFormat = "0." & StrTemp

    If Not (IsNull(rs("InvDate").Value)) Then
                If rs("InvDate").Value = 0 Then
                    SystemOptions.SysInvDateTakeType = InvDateFromLocalCompuer
                ElseIf rs("InvDate").Value = 1 Then
                    SystemOptions.SysInvDateTakeType = InvDateFromLastInvDate
                ElseIf rs("InvDate").Value = 2 Then
                    SystemOptions.SysInvDateTakeType = InvDateFromServerComputer
            
                End If

    Else
              SystemOptions.SysInvDateTakeType = InvDateFromLocalCompuer
    End If

    If Not (IsNull(rs("PurDate").Value)) Then
        If rs("PurDate").Value = 0 Then
            SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer
        ElseIf rs("PurDate").Value = 1 Then
            SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate
        ElseIf rs("PurDate").Value = 2 Then
            SystemOptions.SysPurDateTakeType = InvDateFromServerComputer
        End If

    Else
        SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer
    End If



'*****************************
    If Not (IsNull(rs("CashDate").Value)) Then
                If rs("CashDate").Value = 0 Then
                    SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer
                ElseIf rs("CashDate").Value = 1 Then
                    SystemOptions.SysCashDateTakeType = InvDateFromLastInvDate
                ElseIf rs("CashDate").Value = 2 Then
                    SystemOptions.SysCashDateTakeType = InvDateFromServerComputer
            
                End If

    Else
              SystemOptions.SysCashDateTakeType = InvDateFromLocalCompuer
    End If
    '***************************
    
    

    If Not (IsNull(rs("LockedDate").Value)) Then
 
    Else
 
    End If

'    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
'        SystemOptions.SysCurrentAccountIntervalID = ModAccounts.GetCurrentAccountIntervalID
'    End If

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

    SystemOptions.CLockedDate = IIf(IsNull(rs("LockedDate").Value), "01/01/2050", rs("LockedDate").Value)

    SystemOptions.LockSystem = Val(IIf(IsNull(rs("LockSystem").Value), 0, rs("LockSystem").Value))
    'spareSalimsalimsalim
 
SystemOptions.NoBooking = IIf(IsNull(rs("NoBooking").Value), 0, Val(rs("NoBooking").Value))
If year(Date) > 2018 Then
                    
                     '   If DateDiff("d", CDate(" 20/04/2018"), Date) > 0 Then
                 '   MsgBox "Unable To Connect To the SQL Server .....!", vbCritical
                 '   Cn.Execute "update TblOptions set  Alarm_start='01/01/2000'"
                 '     End
 End If
 'spareSalimsalimsalim
 
   
 
 
 
    If SystemOptions.LockSystem = 10111982 Then
    
      Msg = "SQl Fail To connect "
       MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    On Error GoTo ll
   ' SystemOptions.Alarm_start1 = IIf(IsNull(rs("Alarm_start").value), False, rs("Alarm_start").value)

    Dim X As Integer
    Dim datedifferent As Integer
    Dim startofAlarm As Date

    If Not IsNull(rs("Alarm_start").Value) Then

        startofAlarm = DateAdd("d", -7, rs("Alarm_start").Value)

        datedifferent = DateDiff("d", Date, startofAlarm)

        If datedifferent <= 0 Then
        
            X = DateDiff("d", Date, rs("Alarm_start").Value)

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
    

        

    LoadMainSystemOptions2 = True
    Exit Function
hErr:
    Msg = "Â‰«ﬂ Œÿ« ›Ï Load Main System Options"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    LoadMainSystemOptions2 = False
End Function


