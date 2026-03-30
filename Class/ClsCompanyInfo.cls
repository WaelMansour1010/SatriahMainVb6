VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCompanyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Company_Name_Arabic As String

Private m_Comment_Arabic As String

Private m_Company_Name_Eng As String

Private m_Comment_Eng As String
Private m_VATRegNo As String
Dim m_NoRoudProjectInvoices As Integer
Dim m_VATItems As Integer
Private IntZoomLevel As Integer

Private m_CompanyAddress As String

Private m_CompanyTel As String

Private m_CompanyTax_NO As String

Private m_CompanyTrade_File As String

Private m_CompanyFile_No As String
Private m_ComputerNo As String
Private m_CompanyFax As String
Private m_Company_Mobile As String

Private m_MembershipNo As String
Private m_CompanyEmail As String
Private m_DomainData As String

Private m_CompanyWebSite As String

Private m_CompanyInvComment As String

Private m_CompanyCurrencyUnit As String

Private m_SmallCurrencyUnite As String

Private m_SaleTaxcredit As String

Private m_DefaultPrinterName As String

Private mBranchID As Integer
Public Property Get VATRegNo() As String
    VATRegNo = m_VATRegNo
End Property

Public Property Let SetBranch(ByVal vNewValue As Integer)
     mBranchID = vNewValue
     Class_Initialize
End Property
Public Property Get GetBranch() As String
    GetBranch = mBranchID
End Property
Public Property Get NoRoudProjectInvoices() As String
    NoRoudProjectInvoices = m_NoRoudProjectInvoices
End Property
Public Property Get VATItems() As String
    VATItems = m_VATItems
End Property

Private Sub Class_Initialize()
On Error Resume Next
    Dim RsOpt As ADODB.Recordset
     Dim RsOpt2 As ADODB.Recordset
    Set RsOpt = New ADODB.Recordset
    
    Set RsOpt2 = New ADODB.Recordset
    Dim StrSQL As String
 '   RsOpt.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
     If SystemOptions.WorkWithBranchLogo = False Then
    RsOpt.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
Else
    If mBranchID = 0 Then
            mBranchID = Current_branch
    End If
    If mBranchID = 0 Then
        StrSQL = "SELECT    top 1 *,VATItems = 0  from TblBranchesData "
    Else
        StrSQL = "SELECT     *,VATItems = 0  from TblBranchesData Where (branch_id = " & mBranchID & ")"
    End If
   RsOpt.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

End If
    
    If SystemOptions.WorkWithBranchLogo = False Then
    m_Company_Name_Arabic = IIf(IsNull(RsOpt("Company_Arabic_Name").value), "", RsOpt("Company_Arabic_Name").value)
    m_Comment_Arabic = IIf(IsNull(RsOpt("Company_Comment").value), "", RsOpt("Company_Comment").value)
    m_Company_Name_Eng = IIf(IsNull(RsOpt!Company_Name_Eng), "", RsOpt("Company_Name_Eng").value)
    'm_Comment_Eng = IIf(IsNull(RsOpt("Comment_Eng").Value), "", RsOpt("Comment_Eng").Value)
    m_CompanyAddress = IIf(IsNull(RsOpt("Company_Address").value), "", RsOpt("Company_Address").value)
    m_CompanyTel = IIf(IsNull(RsOpt("Company_Phone").value), "", RsOpt("Company_Phone").value)
    'm_CompanyTax_NO = IIf(IsNull(RsOpt("Tax_No").Value), "", RsOpt("Tax_No").Value)
    'm_CompanyTrade_File = IIf(IsNull(RsOpt("Trade_File").Value), "", RsOpt("Trade_File").Value)
    'm_CompanyFile_No = IIf(IsNull(RsOpt("File_No").Value), "", RsOpt("File_No").Value)
    m_CompanyFax = IIf(IsNull(RsOpt("Fax").value), "", RsOpt("Fax").value)
    m_Company_Mobile = IIf(IsNull(RsOpt("Company_Mobile").value), "", RsOpt("Company_Mobile").value)
    m_CompanyEmail = IIf(IsNull(RsOpt("Company_Maile").value), "", RsOpt("Company_Maile").value)
    m_DomainData = IIf(IsNull(RsOpt("DomainData").value), "", RsOpt("DomainData").value)
    
'    If SystemOptions.VATNoAccordActivity = False Then
    m_VATRegNo = IIf(IsNull(RsOpt!VATRegNo), "", RsOpt("VATRegNo").value)
    m_NoRoudProjectInvoices = IIf(IsNull(RsOpt!NoRoudProjectInvoices), "", RsOpt("NoRoudProjectInvoices").value)
     m_VATItems = IIf(IsNull(RsOpt!VATItems), "", RsOpt("VATItems").value)
'    ElseR
'    m_VATRegNo = GetVATNo(CurrentActivityName)
'    End If

    m_MembershipNo = IIf(IsNull(RsOpt("MembershipNo").value), "", RsOpt("MembershipNo").value)
    m_ComputerNo = IIf(IsNull(RsOpt("ComputerNo").value), "", RsOpt("ComputerNo").value)
    'm_CompanyEmail = IIf(IsNull(RsOpt("email").Value), "", RsOpt("email").Value)
    'm_CompanyWebSite = IIf(IsNull(RsOpt("web_site").Value), "", RsOpt("web_site").Value)
    'm_CompanyInvComment = IIf(IsNull(RsOpt("printed_inv_comment").Value), "", RsOpt("printed_inv_comment").Value)
    'm_CompanyCurrencyUnit = IIf(IsNull(RsOpt("currency_unite").Value), "", RsOpt("currency_unite").Value)
    'm_SmallCurrencyUnite = IIf(IsNull(RsOpt("small_currency_unite").Value), "", RsOpt("small_currency_unite").Value)
    'm_SaleTaxcredit = IIf(IsNull(RsOpt("Sale_Tax_credit").Value), "", RsOpt("Sale_Tax_credit").Value) '
    'm_DefaultPrinterName = IIf(IsNull(RsOpt("DefaultPrinter").Value), "", RsOpt("DefaultPrinter").Value) '
    
    IntZoomLevel = 100 ' IIf(IsNull(RsOpt("RptZoom").Value), 100, RsOpt("RptZoom").Value)     'rs!Comment_Eng
    m_Company_Mobile = IIf(IsNull(RsOpt("Company_Mobile").value), "", RsOpt("Company_Mobile").value)
    
Else


 m_Company_Name_Arabic = IIf(IsNull(RsOpt("branch_name").value), "", RsOpt("branch_name").value)
      m_Company_Name_Eng = IIf(IsNull(RsOpt!branch_namee), "", RsOpt("branch_namee").value)
    m_CompanyTel = IIf(IsNull(RsOpt("Tel").value), "", RsOpt("Tel").value)
     m_Company_Mobile = IIf(IsNull(RsOpt("Tel").value), "", RsOpt("Tel").value)
    m_VATRegNo = IIf(IsNull(RsOpt!VATNO), "", RsOpt("VATNo").value)
    If Trim(m_VATRegNo) = "" Then
        
        
        RsOpt2.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        m_VATRegNo = IIf(IsNull(RsOpt2!VATRegNo), "", RsOpt2("VATRegNo").value)
    End If
  '   m_VATItems = 5
m_CompanyAddress = IIf(IsNull(RsOpt("manger").value), "", RsOpt("manger").value)
     m_VATItems = IIf(IsNull(RsOpt!VATItems), 0, RsOpt("VATItems").value)
End If


    RsOpt.Close
    Set RsOpt = Nothing
End Sub


Public Property Get ArabCompanyName() As String
    ArabCompanyName = m_Company_Name_Arabic
End Property

Public Property Get EngCompanyName() As String
    EngCompanyName = m_Company_Name_Eng
End Property

Public Property Get ArabComment() As String
    ArabComment = m_Comment_Arabic
End Property

Public Property Get EngComment() As String
    EngComment = m_Comment_Eng
End Property

Public Property Get RepZoom() As Integer
    RepZoom = IntZoomLevel
End Property

Public Property Get CompanyAddress() As String
    CompanyAddress = m_CompanyAddress
End Property

Public Property Get CompanyTel() As String
    CompanyTel = m_CompanyTel
End Property

Public Property Get CompanyTax_NO() As String
    CompanyTax_NO = m_CompanyTax_NO
End Property

Public Property Get CompanyTrade_File() As String
    CompanyTrade_File = m_CompanyTrade_File
End Property

Public Property Get CompanyFile_No() As String
    CompanyFile_No = m_CompanyFile_No
End Property

Public Property Get CompanyFax() As String
    CompanyFax = m_CompanyFax
End Property
Public Property Get CompanyMembershipNo() As String
    CompanyMembershipNo = m_MembershipNo
End Property
Public Property Get Company_Mobile() As String
    Company_Mobile = m_Company_Mobile
End Property

Public Property Get ComputerNo() As String
ComputerNo = m_ComputerNo
End Property

Public Property Get CompanyEmail() As String
    CompanyEmail = m_CompanyEmail
End Property

Public Property Get DomainData() As String
    DomainData = m_DomainData
End Property


Public Property Get CompanyWebSite() As String
    CompanyWebSite = m_CompanyWebSite
End Property

Public Property Get CompanyInvComment() As String
    '«· ⁄·Ìﬁ ⁄·Ï ›« Ê—… «·»Ì⁄
    CompanyInvComment = m_CompanyInvComment
End Property

Public Property Get CurrencyUnit() As String
    CurrencyUnit = m_CompanyCurrencyUnit
End Property

Public Property Get SmallCurrencyUnite() As String
    SmallCurrencyUnite = m_SmallCurrencyUnite
End Property

Public Property Get SaleTaxCredit() As String
    '  '—ﬁ„ «·ﬁÌœ ›Ï ÷—Ì»… «·„»Ì⁄« 
    SaleTaxCredit = m_SaleTaxcredit
End Property

Public Property Get DefaultPrinter() As String
    DefaultPrinter = m_DefaultPrinterName
End Property
