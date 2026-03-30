Attribute VB_Name = "salim2014"


Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Dim sql As String

Public Const MF_BYPOSITION = &H400&

Public ReadyToClose As Boolean
'************************Option Explicit********************

Private Const NV_CLOSEMSGBOX = &H5000&
Private Const NV_MOVEMSGBOX = &H5001&
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, _
    ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
    ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mTitle As String
Private mX As Long
Private mY As Long
Private mPause As Long
Private mHandle As Long

  Public cIdentificationID As String
Public cschemeID  As String
Public cStreetName  As String
     Public cAdditionalStreetName As String
         Public cBuildingNumber As String
           Public ccPlotIdentification As String
           
         Public cCityName As String
   Public cPostalZone  As String
    Public cCountrySubentity  As String
     Public cCitySubdivisionName As String
         Public cCountryIdentificationCode As String
           Public cRegistrationName As String
                   Public Companynameforvat As String
             Public cCompanyID As String


Public Function fillmycompanydata(Optional ByVal ActivityTypeIdInv As Integer = 0, Optional ByVal branch_idInv As Integer = 0)
 On Error Resume Next
     '»Ì«‰«   «·‘—ﬂ…
    
     Dim StrRS As New ADODB.Recordset
Dim s As String
    If SystemOptions.ApplyEinvoiceWithActive = True Then
        s = " Select * from tblActivitesType where id = " & IIf(ActivityTypeIdInv = 0, Activity_id, ActivityTypeIdInv)
        StrRS.Open s, Cn, adOpenStatic, adLockReadOnly
    ElseIf SystemOptions.ApplyEinvoiceWithBranch = True Then
        s = " Select Commonname,CSR,Privatekey,SerialNumber,SecretKey,PublickeycertPem,OrganizationName,Invoicetype,DefaultInvoicetype,"
        s = s & " Company_Comment,StreetName,AdditionalStreetName,BuildingNumber,PlotIdentification,CityName,PostalZone,"
        s = s & " CountrySubentity,CitySubdivisionName,Company_Name_Eng,VATRegNo,Company_arabic_Name,industrey,SendingMode "
        s = s & " from TblBranchesData where TblBranchesData.branch_id = " & IIf(branch_idInv = 0, branch_id, branch_idInv)

        StrRS.Open s, Cn, adOpenStatic, adLockReadOnly
    Else
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  End If




If SystemOptions.ApplyEinvoiceWithActive = True Or SystemOptions.ApplyEinvoiceWithBranch = True Then
    SystemOptions.Commonname = IIf(IsNull(StrRS("Commonname").value), -1, StrRS("Commonname").value)
    SystemOptions.SerialNumber = IIf(IsNull(StrRS("SerialNumber").value), -1, StrRS("SerialNumber").value)
    SystemOptions.OrganizationName = IIf(IsNull(StrRS("OrganizationName").value), -1, StrRS("OrganizationName").value)
    SystemOptions.Invoicetype = IIf(IsNull(StrRS("Invoicetype").value), -1, StrRS("Invoicetype").value)
    SystemOptions.DefaultInvoicetype = IIf(IsNull(StrRS("DefaultInvoicetype").value), -1, StrRS("DefaultInvoicetype").value)
    
    SystemOptions.SendingMode = IIf(IsNull(StrRS("SendingMode").value), -1, StrRS("SendingMode").value)
    SystemOptions.industrey = IIf(IsNull(StrRS("industrey").value), -1, StrRS("industrey").value)
    SystemOptions.CSR = IIf(IsNull(StrRS("CSR").value), -1, StrRS("CSR").value)
    SystemOptions.Privatekey = IIf(IsNull(StrRS("Privatekey").value), -1, StrRS("Privatekey").value)
    SystemOptions.PublickeycertPem = IIf(IsNull(StrRS("PublickeycertPem").value), -1, StrRS("PublickeycertPem").value)
    SystemOptions.SecretKey = IIf(IsNull(StrRS("SecretKey").value), -1, StrRS("SecretKey").value)
    
End If

cIdentificationID = IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
cschemeID = "CRN"
cStreetName = IIf(IsNull(StrRS("StreetName").value), "", StrRS("StreetName").value)
cAdditionalStreetName = IIf(IsNull(StrRS("AdditionalStreetName").value), "", StrRS("AdditionalStreetName").value)
cBuildingNumber = IIf(IsNull(StrRS("BuildingNumber").value), "", StrRS("BuildingNumber").value)
ccPlotIdentification = IIf(IsNull(StrRS("PlotIdentification").value), "", StrRS("PlotIdentification").value)
cCityName = IIf(IsNull(StrRS("CityName").value), "", StrRS("CityName").value)
cPostalZone = IIf(IsNull(StrRS("PostalZone").value), "", StrRS("PostalZone").value)
cCountrySubentity = IIf(IsNull(StrRS("CountrySubentity").value), "", StrRS("CountrySubentity").value)
cCitySubdivisionName = IIf(IsNull(StrRS("CitySubdivisionName").value), "", StrRS("CitySubdivisionName").value)
cCountryIdentificationCode = "SA"
 cRegistrationName = IIf(IsNull(StrRS("Company_Name_Eng").value), "", StrRS("Company_Name_Eng").value)
 cCompanyID = IIf(IsNull(StrRS("VATRegNo").value), "", StrRS("VATRegNo").value)
 Companynameforvat = IIf(IsNull(StrRS("Company_arabic_Name").value), "", StrRS("Company_arabic_Name").value)
    
    
    
    
    
End Function


Public Function SENDEINVOICE(Transaction_ID As Long, ISNORMALSALES As Boolean, customerid As Integer, Optional ByVal docType As Integer = 0, Optional ByVal mTableName As String = "transactions", Optional ByVal mFieldIDName As String = "Transaction_ID") As String
 Dim s As String
 Dim ReturnSerial As String
Dim SalesInvoiceDate As Date
 SENDEINVOICE = ""
Dim rsDummy As New ADODB.Recordset
  Dim e As New ClsGLOther
  If docType = 0 Then
    s = " SELECT  dbo.Transactions.DateBaptizing,dbo.Transactions.order_no,  dbo.Transactions.ReturnSerial,  dbo.Transactions.SalesInvoiceDate ,  dbo.Transactions.Transaction_ID, PayeeFinancialAccount =(select IBan  from BanksData  where bankid=Transactions.bankid),     dbo.Transactions.Transaction_ID AS Expr1, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1 AS id, dbo.Transactions.Transaction_Date AS IssueDate, dbo.Transactions.RecTime AS IssueTim, "
      s = s & "                        dbo.Transactions.InvoiceTypeCodeID, dbo.Transactions.InvoiceTypeCodename, dbo.Transactions.DocumentCurrencyCode, dbo.Transactions.TaxCurrencyCode, dbo.Transactions.InvoiceDocumentReferenceID,"
    s = s & "                          dbo.Transactions.AdditionalDocumentReferenceICVUUID, dbo.Transactions.ActualDeliveryDate, dbo.Transactions.LatestDeliveryDate, dbo.Transactions.PaymentMeansCode, dbo.Transactions.InstructionNote,"
    s = s & "                          dbo.Transactions.paymentnote, dbo.TblCustemers.CustGID AS Identificationid, 'CRN' AS schemeID, dbo.TblCustemers.StreetName, dbo.TblCustemers.AdditionalStreetName, dbo.TblCustemers.BuildingNumber,"
    s = s & "                          dbo.TblCustemers.PlotIdentification, dbo.TblCustemers.CityName, dbo.TblCustemers.PostalZone, dbo.TblCustemers.CountrySubentity, dbo.TblCustemers.CitySubdivisionName, dbo.TblCustemers.IdentificationCode,"
    s = s & "                          dbo.TblCustemers.CusNamee AS RegistrationName, dbo.TblCustemers.VATNO AS CompanyID, dbo.Transactions.LblDiscountsTotal AS allowancechargeAmount, 'Discount' AS AllowanceChargeReason, 'S' AS TaxCategoryID,"
    s = s & "                          '15' AS TaxCategoryPercent, dbo.Transactions.last_changed, dbo.Transactions.Transaction_NetValue AS PayableAmount, dbo.Transactions.AdvPay AS PrepaidAmount, dbo.transactionsVatDetails.SingedXMLFileName,"
    s = s & "                          dbo.transactionsVatDetails.PIH, dbo.transactionsVatDetails.QRCode, dbo.transactionsVatDetails.UUID, dbo.transactionsVatDetails.InvoiceHash, dbo.transactionsVatDetails.EncodedInvoice,"
    s = s & "                          dbo.transactionsVatDetails.SingedXML,  dbo.transactionsVatDetails.QrCodeDataPath"
    s = s & "  FROM            dbo.TblCustemers INNER JOIN"
    s = s & "                          dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID left OUTER JOIN"
    s = s & "                          dbo.transactionsVatDetails ON dbo.Transactions.Transaction_ID = dbo.transactionsVatDetails.Transaction_ID and isnull(transactionsVatDetails.isdeleted,0)=0  and isNull(transactionsVatDetails.DocType,0) = 0"


    s = s & "  Where   dbo.Transactions.Transaction_ID =" & Transaction_ID
ElseIf docType = 1 Then
        s = " SELECT  dbo.project_billl.bill_date DateBaptizing,dbo.Transactions.order_no,  dbo.Transactions.ReturnSerial,  dbo.Transactions.SalesInvoiceDate ,  dbo.Transactions.Transaction_ID, PayeeFinancialAccount =(select IBan  from BanksData  where bankid=Transactions.bankid),     dbo.Transactions.Transaction_ID AS Expr1, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1 AS id, dbo.Transactions.Transaction_Date AS IssueDate, dbo.Transactions.RecTime AS IssueTim, "
   
   
    s = " SELECT"
    s = s & " dbo.project_billl.bill_date DateBaptizing"
    s = s & " ,dbo.project_billl.order_no"
    s = s & " ,0 ReturnSerial"
    s = s & " ,dbo.project_billl.bill_date SalesInvoiceDate"
    s = s & " ,dbo.project_billl.id Transaction_ID"
    s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "         WHERE bankid = 0)"
    s = s & "    ,dbo.project_billl.ID AS Expr1"
    s = s & " ,1 as Transaction_Type"
    s = s & " ,dbo.project_billl.NoteSerial1 AS id"
    s = s & "    ,dbo.project_billl.bill_date AS IssueDate"
    s = s & "    ,dbo.project_billl.RecTime AS IssueTim"
   s = s & " ,dbo.project_billl.InvoiceTypeCodeID"
   s = s & " ,dbo.project_billl.InvoiceTypeCodename"
   s = s & " ,dbo.project_billl.DocumentCurrencyCode"
   s = s & " ,dbo.project_billl.TaxCurrencyCode"
   s = s & " ,dbo.project_billl.InvoiceDocumentReferenceID"
   s = s & " ,dbo.project_billl.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.project_billl.ActualDeliveryDate"
   s = s & " ,dbo.project_billl.LatestDeliveryDate"
   s = s & " ,dbo.project_billl.PaymentMeansCode"
   s = s & " ,dbo.project_billl.InstructionNote"
   s = s & " ,dbo.project_billl.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID"
   s = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,1 last_changed"
   s = s & " ,dbo.project_billl.TotalValue AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
    s = s & " From dbo.TblCustemers"
    s = s & " INNER JOIN dbo.project_billl"
    s = s & " inner join projects On project_no = projects.id"
    s = s & "     ON dbo.TblCustemers.CusID = dbo.projects.End_user_id"
s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
s = s & "     ON dbo.project_billl.ID = dbo.transactionsVatDetails.Transaction_ID"
s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
        


    s = s & "  Where   dbo.project_billl.ID =" & Transaction_ID


ElseIf docType = 3 Then
     s = " SELECT"
    s = s & " Invoicetype"
   s = s & " ,dbo.Notes.ErrorMessageS"
   s = s & " ,dbo.Notes.NoteDate DateBaptizing"
   s = s & " ,cast (dbo.Notes.order_no    as VARCHAR(10)) order_no"
   s = s & " ,0 ReturnSerial"
   s = s & " ,dbo.Notes.NoteDate SalesInvoiceDate"
   s = s & " ,dbo.Notes.Noteid Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "         WHERE bankid = 0)"
   s = s & " ,dbo.Notes.Noteid AS Expr1"
   s = s & " ,381 as Transaction_Type = (case Notes.NoteType when 9082 then 383 when 9083 then 381 end )"
   s = s & " ,dbo.Notes.NoteSerial1 AS id"
   s = s & " ,dbo.Notes.bill_date AS IssueDate"
   s = s & " ,dbo.Notes.RecTime AS IssueTim"
   s = s & " ,dbo.Notes.InvoiceTypeCodeID"
   s = s & " ,dbo.Notes.InvoiceTypeCodename"
   s = s & " ,dbo.Notes.DocumentCurrencyCode"
   s = s & " ,dbo.Notes.TaxCurrencyCode"
   s = s & " ,dbo.Notes.InvoiceDocumentReferenceID"
   s = s & " ,dbo.Notes.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.Notes.ActualDeliveryDate"
   s = s & " ,dbo.Notes.LatestDeliveryDate"
   s = s & " ,dbo.Notes.PaymentMeansCode"
   s = s & " ,dbo.Notes.InstructionNote"
   s = s & " ,dbo.Notes.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID"
   
   's = s & " ,Notes.DiscountGMater + Notes.Discount4 + Notes.advancedPayment AS allowancechargeAmount"
    's = s & " ,Notes.discount + Notes.DiscountGMater +  Notes.advancedPayment AS allowancechargeAmount"
    s = s & " ,0  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,Notes.last_changed"
   
   's = s & " ,Notes.total+ Notes.FATValue AS PayableAmount"
   s = s & " ,Notes.TotalValue AS PayableAmount"
   's = s & " ,Notes.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,3 AS DocType"

'«‘⁄«— „œÌ‰   =9082
'«‘⁄«— œ«∆‰ = 9083
    s = s & " From dbo.TblCustemers"
    s = s & " INNER JOIN dbo.Notes"
    's = s & " inner join projects On project_no = projects.id"
    s = s & " ON dbo.TblCustemers.CusID = dbo.projects.End_user_id"
    s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.Notes.NoteID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"

    s = s & " Where 1 = 1 and Notes.NoteType In (9083,9082)"
    
    


    s = s & "  Where   dbo.Notes.NoteID =" & Transaction_ID


End If




 Set rsDummy = New ADODB.Recordset
    
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
 
      
 
  e.ID = IIf(IsNull(rsDummy("id").value), "", rsDummy("id").value)
  e.IssueDate = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
  e.IssueTim = IIf(IsNull(rsDummy("IssueTim").value), "", rsDummy("IssueTim").value)
 e.InvoiceTypeCodeID = IIf(IsNull(rsDummy("InvoiceTypeCodeID").value), "", rsDummy("InvoiceTypeCodeID").value)
e.InvoiceTypeCodename = IIf(IsNull(rsDummy("InvoiceTypeCodename").value), "", rsDummy("InvoiceTypeCodename").value)
If customerid = 1 Or customerid = 2 Then
e.InvoiceTypeCodename = "0200000"
End If

    

e.DocumentCurrencyCode = IIf(IsNull(rsDummy("DocumentCurrencyCode").value), "", rsDummy("DocumentCurrencyCode").value)
e.TaxCurrencyCode = IIf(IsNull(rsDummy("TaxCurrencyCode").value), "", rsDummy("TaxCurrencyCode").value)
'e.InvoiceDocumentReferenceID = IIf(IsNull(rsDummy("InvoiceDocumentReferenceID").value), "", rsDummy("InvoiceDocumentReferenceID").value)
e.docType = "38801"
If val(e.InvoiceTypeCodeID) = 381 Then '«‘⁄«— œ«∆‰ „—œÊœ« 
e.docType = "38101"
ReturnSerial = IIf(IsNull(rsDummy("ReturnSerial").value), "", rsDummy("ReturnSerial").value)
SalesInvoiceDate = IIf(IsNull(rsDummy("SalesInvoiceDate").value), "", rsDummy("SalesInvoiceDate").value)
e.InvoiceDocumentReferenceID = "?Invoice Number: " & ReturnSerial & "; Invoice Issue Date: " & Format(SalesInvoiceDate, "yyyy-mm-dd") & "?"

End If

      
If val(e.InvoiceTypeCodeID) = 383 Then ' «‘⁄«— „œÌ‰ “Ì  „»Ì⁄« 
e.docType = "38301"
ReturnSerial = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
SalesInvoiceDate = IIf(IsNull(rsDummy("DateBaptizing").value), "", rsDummy("DateBaptizing").value)
e.InvoiceDocumentReferenceID = "?Invoice Number: " & ReturnSerial & "; Invoice Issue Date: " & Format(SalesInvoiceDate, "yyyy-mm-dd") & "?"
End If


e.AdditionalDocumentReferenceICVUUID = IIf(IsNull(rsDummy("AdditionalDocumentReferenceICVUUID").value), "", rsDummy("AdditionalDocumentReferenceICVUUID").value)
e.ActualDeliveryDate = IIf(IsNull(rsDummy("ActualDeliveryDate").value), "", rsDummy("ActualDeliveryDate").value)
e.LatestDeliveryDate = IIf(IsNull(rsDummy("LatestDeliveryDate").value), "", rsDummy("LatestDeliveryDate").value)
e.PaymentMeansCode = IIf(IsNull(rsDummy("PaymentMeansCode").value), "", rsDummy("PaymentMeansCode").value)
e.InstructionNote = IIf(IsNull(rsDummy("InstructionNote").value), "", rsDummy("InstructionNote").value)
e.PayeeFinancialAccount = IIf(IsNull(rsDummy("PayeeFinancialAccount").value), "", rsDummy("PayeeFinancialAccount").value)
e.paymentnote = IIf(IsNull(rsDummy("paymentnote").value), "", rsDummy("paymentnote").value)
e.Identificationid = IIf(IsNull(rsDummy("Identificationid").value), "", rsDummy("Identificationid").value)
e.schemeID = IIf(IsNull(rsDummy("schemeID").value), "", rsDummy("schemeID").value)
e.StreetName = IIf(IsNull(rsDummy("StreetName").value), "", rsDummy("StreetName").value)
e.AdditionalStreetName = IIf(IsNull(rsDummy("AdditionalStreetName").value), "", rsDummy("AdditionalStreetName").value)
e.BuildingNumber = IIf(IsNull(rsDummy("BuildingNumber").value), "", rsDummy("BuildingNumber").value)
e.PlotIdentification = IIf(IsNull(rsDummy("PlotIdentification").value), "", rsDummy("PlotIdentification").value)
e.CityName = IIf(IsNull(rsDummy("CityName").value), "", rsDummy("CityName").value)
e.PostalZone = IIf(IsNull(rsDummy("PostalZone").value), "", rsDummy("PostalZone").value)
e.CountrySubentity = IIf(IsNull(rsDummy("CountrySubentity").value), "", rsDummy("CountrySubentity").value)
e.CitySubdivisionName = IIf(IsNull(rsDummy("CitySubdivisionName").value), "", rsDummy("CitySubdivisionName").value)
e.IdentificationCode = IIf(IsNull(rsDummy("IdentificationCode").value), "", rsDummy("IdentificationCode").value)
e.RegistrationName = IIf(IsNull(rsDummy("RegistrationName").value), "", rsDummy("RegistrationName").value)
e.CompanyID = IIf(IsNull(rsDummy("CompanyID").value), "", rsDummy("CompanyID").value)
e.allowancechargeAmount = IIf(IsNull(rsDummy("allowancechargeAmount").value), 0, rsDummy("allowancechargeAmount").value)
e.AllowanceChargeReason = IIf(IsNull(rsDummy("AllowanceChargeReason").value), "", rsDummy("AllowanceChargeReason").value)
e.TaxCategoryID = IIf(IsNull(rsDummy("TaxCategoryID").value), "", rsDummy("TaxCategoryID").value)
e.TaxCategoryPercent = IIf(IsNull(rsDummy("TaxCategoryPercent").value), "", rsDummy("TaxCategoryPercent").value)
e.PayableAmount = IIf(IsNull(rsDummy("PayableAmount").value), "", rsDummy("PayableAmount").value)
e.PrepaidAmount = IIf(IsNull(rsDummy("PrepaidAmount").value), 0, rsDummy("PrepaidAmount").value)
  e.Transaction_ID = IIf(IsNull(rsDummy("Transaction_ID").value), "", rsDummy("Transaction_ID").value)
  e.InvoiceHash = IIf(IsNull(rsDummy("InvoiceHash").value), "", rsDummy("InvoiceHash").value)
   e.SingedXML = IIf(IsNull(rsDummy("SingedXML").value), "", rsDummy("SingedXML").value)
   e.EncodedInvoice = IIf(IsNull(rsDummy("EncodedInvoice").value), "", rsDummy("EncodedInvoice").value)
   e.UUID = IIf(IsNull(rsDummy("UUID").value), "", rsDummy("UUID").value)
   e.QRCode = IIf(IsNull(rsDummy("QRCode").value), "", rsDummy("QRCode").value)
   e.PIH = IIf(IsNull(rsDummy("PIH").value), "", rsDummy("PIH").value)
   e.SingedXMLFileName = IIf(IsNull(rsDummy("SingedXMLFileName").value), "", rsDummy("SingedXMLFileName").value)
e.QrCodeDataPath = IIf(IsNull(rsDummy("QrCodeDataPath").value), "", rsDummy("QrCodeDataPath").value)
e.generateInvoice
  SENDEINVOICE = e.ErrorMessageS
End Function


Public Function MsgBoxMove(ByVal hWnd As Long, ByVal inPrompt As String, _
        ByVal inTitle As String, ByVal inButtons As Long, _
               ByVal inX As Long, ByVal inY As Long) As Integer
     mTitle = inTitle: mX = inX:  mY = inY
     SetTimer hWnd, NV_MOVEMSGBOX, 0&, AddressOf NewTimerProc
     MsgBoxMove = MessageBox(hWnd, inPrompt, inTitle, inButtons)
End Function


Public Function MsgBoxPause(ByVal hWnd As Long, ByVal inPrompt As String, _
        ByVal inTitle As String, ByVal inButtons As Long, _
        ByVal inPause As Integer) As Integer
     mTitle = inTitle: mPause = inPause * 1000: mX = 0:  mY = 0
     SetTimer hWnd, NV_MOVEMSGBOX, 0&, AddressOf NewTimerProc
     SetTimer hWnd, NV_CLOSEMSGBOX, mPause, AddressOf NewTimerProc
     MsgBoxPause = MessageBox(hWnd, inPrompt, inTitle, inButtons)
End Function


Public Function NewTimerProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
    KillTimer hWnd, wParam
    Select Case wParam
         Case NV_CLOSEMSGBOX
              ' A system class is a window class registered by the system which cannot
              ' be destroyed by a processed, e.g. #32768 (a menu), #32769 (desktop
              ' window), #32770 (dialog box), #32771 (task switch window).
             mHandle = FindWindow("#32770", mTitle)
             If mHandle <> 0 Then
                  SetForegroundWindow mHandle
                  Sendkeys "{enter}"
             End If
             
        Case NV_MOVEMSGBOX
             mHandle = FindWindow("#32770", mTitle)
             If mHandle <> 0 Then
                  Dim W As Single, H As Single
                  Dim mBox As RECT
                  W = Screen.Width / Screen.TwipsPerPixelX
                  H = Screen.Height / Screen.TwipsPerPixelY
                  GetWindowRect mHandle, mBox
                  If mX > (W - (mBox.right - mBox.left) - 1) Then mX = (W - (mBox.right - mBox.left) - 1)
                  If mY > (H - (mBox.bottom - mBox.top) - 1) Then mY = (H - (mBox.bottom - mBox.top) - 1)
                  If mX < 1 Then mX = 1: If mY < 1 Then mY = 1
                    ' SWP_NOSIZE is to use current size, ignoring 3rd & 4th parameters.
                  SetWindowPos mHandle, HWND_TOPMOST, mX, mY, 0, 0, SWP_NOSIZE
             End If
    End Select
End Function



Public Function GetQtySqlYearly(LngItemID As Long, _
                                     Optional LngStoreID As Long = 0, _
                                     Optional YearID As Integer, _
                                     Optional LngColorID As Long = 1, _
                                     Optional StrItemSize As String = "", _
                                     Optional ClassId As Long = 1) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(dbo.Transaction_Details.Quantity ) AS totalqty"
sql = sql & " FROM         dbo.Transaction_Details INNER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
sql = sql & "                     dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql & " Where (dbo.transactions.StoreID = " & LngStoreID & ") And "
sql = sql & "                       (dbo.Transactions.Transaction_Type = 21) and (YEAR(Transactions.Transaction_Date) = " & YearID & ")"
sql = sql & " GROUP BY dbo.Transaction_Details.Item_ID"
sql = sql & " Having (dbo.Transaction_Details.Item_ID = " & LngItemID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetQtySqlYearly = IIf(IsNull(rs2("totalqty").value), 0, rs2("totalqty").value)
Else
GetQtySqlYearly = 0
End If

End Function
 Public Function SumDay(Optional Fild As String, Optional EmpID As Integer) As Double
Dim Filedname As String
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Filedname = "dbo.TbLSheft." & Fild
sql = " SELECT     SUM(" & Filedname & ") AS SmSHift"
sql = sql & " FROM         dbo.TbLSheft LEFT OUTER JOIN"
sql = sql & "                      dbo.TblShiftWorker ON dbo.TbLSheft.SeftCode = dbo.TblShiftWorker.ShiftID"
sql = sql & " GROUP BY dbo.TblShiftWorker.EmpID"
sql = sql & " Having (dbo.TblShiftWorker.EmpID <> 0) and (dbo.TblShiftWorker.EmpID = " & EmpID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
SumDay = IIf(IsNull(Rs3("SmSHift").value), 0, Rs3("SmSHift").value)
Else
SumDay = 0
End If
End Function


Public Function SumHour(Optional EmpID1 As Integer, Optional NoDay As Integer) As Double
Select Case NoDay
Case 7
SumHour = SumDay("NoHSat", EmpID1)
Case 6
SumHour = SumDay("NoHSun", EmpID1)
Case 5
SumHour = SumDay("NoMon", EmpID1)
Case 4
SumHour = SumDay("NoHTues", EmpID1)
Case 3
SumHour = SumDay("NoHWed", EmpID1)
Case 2
SumHour = SumDay("NoHThru", EmpID1)
Case 1
SumHour = SumDay("NoHFri", EmpID1)
End Select
End Function

Public Function SumDayVaction(Optional EmID As Integer, Optional MonthID As Integer)
Dim sql As String
Dim i As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT    [Date] "
sql = sql & " From dbo.TblVacationschedule22"
sql = sql & " Where (ISVac = 1) And (Month([Date]) = " & MonthID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
SumDayVaction = 0
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
SumDayVaction = SumDayVaction + SumHour(EmID, Weekday(IIf(IsNull(Rs3("Date").value), Date, Rs3("Date").value)))
Rs3.MoveNext
Next i
Else
SumDayVaction = 0
End If
End Function

   
 
 
 
 Public Function AqarCommisionType(Aqarid As Double, Optional AmolaValus As Double, Optional ownerid As Double) As Integer
'If Aqarid <> 0 Then
Dim Rs9  As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
sql = "select * from tblaqar where Aqarid =" & Aqarid & ""
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
 
                      If Not IsNull(Rs9("TypAmola").value) Then
                     
                     AqarCommisionType = IIf(IsNull(Rs9("TypAmola").value), 0, Rs9("TypAmola").value)
                     AmolaValus = IIf(IsNull(Rs9("AmolaValus").value), 0, Rs9("AmolaValus").value)
                     ownerid = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)
                    Else
                   ownerid = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)
                    AmolaValus = 0
                    ownerid = ownerid
                    AmolaValus = 0
                     End If
                     
 End If
End Function
 
Public Function PrintSimpleReport(sql As String, path As String, Optional StrWhere As String = "", Optional X As String = "")
 
    Dim rs As New ADODB.Recordset
 Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report

     rs.Open sql & " " & StrWhere, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(path)
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    If X <> "" Then
     xReport.ParameterFields(1).AddCurrentValue X
    End If
    
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtPath = (path)
    FrmReport.CRViewer.viewReport
    FrmReport.show
 
    Screen.MousePointer = vbDefault
 
    Sendkeys "{RIGHT}"
    
End Function

Public Function GetQuepEmpVocation(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Integer, _
                                      Optional Emp_id1 As Integer = 0, _
                                      Optional ByRef EmpCode1 As String, _
                                      Optional ByRef Name1 As String, _
                                      Optional ByRef Name As String, _
                                      Optional ByRef Mobile As String, _
                                       Optional ByRef phone As String, _
                                        Optional ByRef boxmail As String, _
                                       Optional ByRef fax As String, _
                                       Optional ByRef mail As String, _
                                       Optional ByRef adress As String, _
                                        Optional ByRef ZipCode As String, _
                                       Optional ByRef DigCus As String, Optional SpecificHolidyaType1 As Boolean, Optional SpecificHolidyaType2 As Boolean)
                                    
            'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblQuesEmp where id= " & Emp_id1 & " and (HolidayType=1)"
    

 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'Dim name, mobile, phone, boxmail, fax, mail, adress, ZipCode, DigCus As String
    If rs.RecordCount > 0 Then
        Emp_id = val(IIf(IsNull(rs("EmpID").value), 0, rs("EmpID").value))
     '   EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
 'name = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
 'mobile = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
 'phone = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
 'boxmail = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
 'fax = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
 'mail = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
 'adress = IIf(IsNull(rs("Address").value), "", rs("Address").value)
 'ZipCode = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
 'DigCus = IIf(IsNull(rs("TypeCustomer").value), "", rs("TypeCustomer").value)
 If (rs("SpecificHolidyaType1").value = True) Then
 SpecificHolidyaType1 = True
 Else
  SpecificHolidyaType1 = False
  End If
   If (rs("SpecificHolidyaType2").value = True) Then
 SpecificHolidyaType2 = True
 Else
  SpecificHolidyaType2 = False
  End If
    Else
        Emp_id = 0
    End If

    rs.Close
End If
End Function
 
Public Sub RemoveMenus(Frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
On Error Resume Next
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(Frm.hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

  Public Function CreateNotes(ByRef NoteID As Long, NoteDate As Date, branch_no As Integer, NoteType As Integer, Note_Value As Double _
 , Optional ByRef NoteSerial As String, Optional NoteSerial1 As String, Optional tablename As String, Optional Filedname As String, Optional Filedvalue As Long, Optional Remark As String, Optional NoteDateH As String, Optional ManualNO As String, Optional NoteIDFiled As String, Optional SerialFiled As String, Optional installIDCont As String, Optional isOpenBalance As Boolean = False, Optional TblLCID As Long = 0, Optional RowID As String = "", Optional RowIDField As String = "", Optional ByVal IsHiddenInv As Boolean = False)
 Dim StrSQL As String
 Dim RsNotesGeneral As New ADODB.Recordset
    'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   
    If isOpenBalance Then
        StrSQL = "SELECT     dbo.Notes1.* from dbo.Notes1 Where (NoteID = -1)"
    Else
        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    End If
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
   RsNotesGeneral.AddNew
   If isOpenBalance Then
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes1", "NoteID", "", True))
        
        'NoteSerial1 = OpeningVoucher_coding(val(branch_no), NoteDate, 3, 101)
   Else
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    End If
    NoteID = RsNotesGeneral("NoteID").value
   RsNotesGeneral.update
     
     If NoteSerial = "" Or NoteSerial = "0" Then
        If isOpenBalance Then
            NoteSerial = OpeningVoucher_coding(val(my_branch), NoteDate, 3, 101)
        Else
            NoteSerial = Notes_coding(val(branch_no), NoteDate)
          End If
         
       End If
 
        RsNotesGeneral("NoteSerial").value = IIf(Trim(NoteSerial) = "", Null, Trim(NoteSerial))
        If isOpenBalance Then
            NoteSerial1 = NoteSerial
        End If
    If IsNumeric(NoteSerial1) Then
      RsNotesGeneral("NoteSerial1").value = IIf(Trim(NoteSerial1) = "", Null, Trim(NoteSerial1))
     Else
     RsNotesGeneral("NoteSerial1").value = Null
     End If
     
      If isOpenBalance Then
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(NoteSerial1) = "", Null, Trim(NoteSerial1))
      End If
    RsNotesGeneral.update
    RsNotesGeneral("NoteDate").value = NoteDate
    RsNotesGeneral("NoteDateH").value = NoteDateH
    
     If IsHiddenInv Then
        RsNotesGeneral("IsHiddenInv").value = 1
    Else
        RsNotesGeneral("IsHiddenInv").value = 0
    End If
    
  
  
    RsNotesGeneral("NoteType").value = NoteType
    RsNotesGeneral("Note_Value").value = Note_Value
    
    If RowID <> "" Then
        'RsNotesGeneral("RowID").value = "{" & Trim(RowID) & "}"
        
            If InStr(RowID, "{") Then
    Else
        RowID = "{" & RowID & "}"
    End If
        RsNotesGeneral("RowID").value = Trim(RowID)
    End If
    If Not isOpenBalance Then
        RsNotesGeneral("installIDCont").value = val(installIDCont)
        RsNotesGeneral("ManualNo").value = ManualNO
    Else
    
    End If
  '  RsNotesGeneral("ManualNo").value = ManualNO
    RsNotesGeneral("TblLCID").value = TblLCID
   

   RsNotesGeneral("remark").value = Remark
 
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '??I C???I
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(NoteType) '  ??I C??C?
    RsNotesGeneral("sanad_year").value = year(NoteDate)
    RsNotesGeneral("sanad_month").value = Month(NoteDate)
    RsNotesGeneral("branch_no").value = branch_no
    RsNotesGeneral("note_value_by_characters").value = WriteNo(val(Note_Value), 0, True)
    RsNotesGeneral.update
    RsNotesGeneral.Close
  If SerialFiled = "" Then
    SerialFiled = "NoteSerial"
  End If
 
If tablename <> "" Then
If NoteIDFiled = "" Then
   StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
   Else
   StrSQL = "update  " & tablename & "   set " & NoteIDFiled & "=" & NoteID & "," & SerialFiled & "='" & NoteSerial & "'"
End If
If RowIDField <> "" Then

    StrSQL = StrSQL & " Where " & RowIDField & "   = '" & RowID & "'"
Else
    StrSQL = StrSQL & " Where " & Filedname & " = " & Filedvalue & " "
End If
  
  If TblLCID <> 0 Then
    StrSQL = StrSQL & " and TblLCID = " & TblLCID
  End If
    Cn.Execute StrSQL
 
 End If
End Function

    


 
'waelGrid

' œÏ   Õÿ ›Ï «Ï „ÊœÌÊ· ⁄«„
Public Sub GridKeyDown(G As Object, _
                       ByRef KeyCode As Integer, _
                       Shift As Integer, _
                       Optional ByVal CancelInsert As Boolean, _
                       Optional ByVal CancelDelete As Boolean, _
                       Optional SerStartRow As Single = 1, _
                        Optional colStartRow As String = "")
            
        Dim FirstVisibleCol As Long
        Dim LastVisibleCol As Long
        Dim FirstVisibleRow As Long
        Dim LastVisibleRow As Long
        Dim NextVisibleCol As Long
        Dim NextVisibleRow As Long
        
        
        Dim X As Integer
        
         If G.Row <= 0 Or G.Col < 0 Then Exit Sub
    ' ************************
    If SerStartRow < G.FixedRows + 1 Then
        SerStartRow = G.FixedRows
    End If
    ' ************************
    FirstVisibleRow = 0
    For X = G.FixedRows To G.rows - 1
        If Not G.RowHidden(X) Then FirstVisibleRow = X: Exit For
    Next
    ' ************************
    LastVisibleRow = 0
    For X = G.rows - 1 To G.FixedRows Step -1
        If Not G.RowHidden(X) Then LastVisibleRow = X: Exit For
    Next
    ' ************************
    FirstVisibleCol = 0
    For X = 1 To G.Cols - 1
        If Not G.ColHidden(X) Then FirstVisibleCol = X: Exit For
    Next
    ' ************************
    LastVisibleCol = 0
    For X = G.Cols - 1 To 1 Step -1
        If Not G.ColHidden(X) Then LastVisibleCol = X: Exit For
    Next
            Select Case KeyCode

            Case vbKeyDelete
                If Shift = 1 Then
                    If CancelDelete Then Exit Sub
                    ' ***************************
                    If LastVisibleRow >= SerStartRow + 1 Then    'G.Rows >= 3 Then
                        If G.Row < LastVisibleRow Then    'G.Rows - 1 Then
                            G.RemoveItem G.Row
                        Else
                            G.RemoveItem G.Row
                            ' ************************
                            LastVisibleRow = 0
                            For X = G.rows - 1 To 1 Step -1
                                If Not G.RowHidden(X) Then LastVisibleRow = X: Exit For
                            Next
                            ' ************************
                            G.Row = LastVisibleRow    'G.Rows - 1
                        End If
                        GridSerial G, , SerStartRow
                    Else
                        G.cell(flexcpText, SerStartRow, 1, SerStartRow, G.Cols - 1) = ""
                    End If
                    GridSerial G, , SerStartRow
                End If
            Case vbKeyReturn
                If G.Row = LastVisibleRow Then    'G.Rows - 1 Then ' BottomRow Then
                    If colStartRow <> "" Then
        
                        If LastVisibleCol = G.ColIndex(colStartRow) Then
                            FirstVisibleCol = G.ColIndex(colStartRow)
                        End If
                    End If

                    If G.Col = LastVisibleCol Then    'G.Cols - 1 Then
        
                        If CancelInsert Then G.FinishEditing False: GoTo mExit
                        ' ***************************
                        G.AddItem "", G.rows
                        ' *****************
                        GridSerial G, , SerStartRow
                        G.Row = G.rows - 1
                        G.Col = FirstVisibleCol
                    Else
                        NextVisibleCol = G.Col
                        For X = G.Col + 1 To G.Cols - 1
                            If Not G.ColHidden(X) Then NextVisibleCol = X: Exit For
                        Next
                        ' ************************
                        G.Col = NextVisibleCol
                    End If
            Else
                If G.Col = LastVisibleCol Then    'G.Cols - 1 Then
                    G.Col = FirstVisibleCol    '1
                    NextVisibleRow = G.Row
                    For X = G.Row + 1 To G.rows - 1
                        If Not G.RowHidden(X) Then NextVisibleRow = X: Exit For
                    Next
                    ' ************************
                    G.Row = NextVisibleRow
                Else
                    NextVisibleCol = G.Col
                    For X = G.Col + 1 To G.Cols - 1
                        If Not G.ColHidden(X) Then NextVisibleCol = X: Exit For
                    Next
                    ' ************************
                    G.Col = NextVisibleCol
                End If
            End If
mExit:
        KeyCode = 0
        End Select

        Exit Sub
End Sub
Public Sub GridSerial(vsGrd As Object, _
                      Optional chkRowHidden As Boolean = False, _
                      Optional SerStartRow As Single = 0)
    Dim j As Long
    Dim ss As Long
    If SerStartRow = 0 Then SerStartRow = vsGrd.FixedRows

    For ss = SerStartRow To vsGrd.rows - 1
        If chkRowHidden Then
            If Not vsGrd.RowHidden(ss) Then j = j + 1
        Else
            j = j + 1
        End If
        vsGrd.TextMatrix(ss, 0) = j
    Next

End Sub





 
 
 
 

 Public Sub loadgrid(ByVal Sqlstmt As String, _
                          ByRef tGrd As Control, _
                          Optional ResetRows As Boolean = True, _
                          Optional InsertRow As Boolean = False, _
                          Optional mReCreateColumns As Boolean = False, Optional ByVal Conn As ADODB.Connection = Nothing)
    Dim tRs As New ADODB.Recordset
  
 
     If Conn Is Nothing Then
        tRs.Open Sqlstmt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        tRs.Open Sqlstmt, Conn, adOpenKeyset, adLockReadOnly
    End If
  
    ' ******************************************
    If ResetRows Then tGrd.rows = tGrd.FixedRows
    ' ******************************************
    Dim i As Long
    If mReCreateColumns Then
        tGrd.Cols = 1
        tGrd.Cols = tRs.Fields.count + 1
        For i = 1 To tGrd.Cols - 1
            tGrd.ColKey(i) = tRs.Fields.Item(i - 1).Name
            tGrd.TextMatrix(0, i) = tRs.Fields.Item(i - 1).Name
        Next
    End If
    ' ******************************************
    ' ******************************************
    tGrd.Redraw = flexRDNone
    ' ******************************************
    i = tGrd.rows
    Dim sCur As Long, j As Long
    sCur = 0
    Do While Not tRs.EOF
        tGrd.AddItem i - tGrd.FixedRows + 1
        For j = 0 To tRs.Fields.count - 1
            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
                If tRs.Fields.Item(j).type = adCurrency Then
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = (val(tRs.Fields.Item(j).value & ""))
                Else
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = Trim(tRs.Fields.Item(j).value & "")
                End If
            End If
        Next
        i = i + 1
        sCur = sCur + 1

        tRs.MoveNext
    Loop
    tRs.Close
    Set tRs = Nothing

    If InsertRow Then tGrd.AddItem tGrd.rows - tGrd.FixedRows + 1
    tGrd.Redraw = flexRDDirect
End Sub




Public Sub loadgridRS(ByVal tRs As ADODB.Recordset, _
                      ByRef tGrd As Control, _
                      Optional ResetRows As Boolean = True, _
                      Optional InsertRow As Boolean = False, _
                      Optional ByVal mReCreateColumns As Boolean = False)

    Dim i As Long, j As Long

    If tRs Is Nothing Then Exit Sub

    ' ›÷Ì «·’›Ê›
    If ResetRows Then tGrd.rows = tGrd.FixedRows

    ' ·Ê ⁄«Ì“  ⁄Ìœ «·√⁄„œ… (“Ì „« »‰⁄„· ›Ì «·” Ê—œ)
    If mReCreateColumns Then
        tGrd.Cols = 1
        tGrd.Cols = tRs.Fields.count + 1
        For i = 1 To tGrd.Cols - 1
            tGrd.ColKey(i) = tRs.Fields(i - 1).Name
            tGrd.TextMatrix(0, i) = tRs.Fields(i - 1).Name
        Next i
    End If

    tGrd.Redraw = flexRDNone

    i = tGrd.rows
    Do While Not tRs.EOF
        tGrd.AddItem i - tGrd.FixedRows + 1
        For j = 0 To tRs.Fields.count - 1
            If tGrd.ColIndex(tRs.Fields(j).Name) <> -1 Then
                If tRs.Fields(j).type = adCurrency Then
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields(j).Name)) = _
                        val(tRs.Fields(j).value & "")
                Else
                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields(j).Name)) = _
                        Trim$(tRs.Fields(j).value & "")
                End If
            End If
        Next j
        i = i + 1
        tRs.MoveNext
    Loop

    If InsertRow Then tGrd.AddItem tGrd.rows - tGrd.FixedRows + 1

    tGrd.Redraw = flexRDDirect
End Sub




'-----------------------------------------
'  ⁄»∆… «·Ã—Ìœ „‰ Recordset Ã«Â“
'-----------------------------------------
Public Sub loadgridRS2(ByRef rs As ADODB.Recordset, _
                      ByRef tGrd As Control, _
                      Optional ResetRows As Boolean = True, _
                      Optional InsertRow As Boolean = False, _
                      Optional ByVal ReCreateColumns As Boolean = False)

    Dim i As Long, j As Long

    If rs Is Nothing Then Exit Sub
    If rs.State = adStateClosed Then Exit Sub

    ' «„”Õ «·’›Ê›
    If ResetRows Then tGrd.rows = tGrd.FixedRows

    ' ·Ê ⁄«Ì“  ⁄Ìœ ≈‰‘«¡ «·√⁄„œ…
    If ReCreateColumns Then
        tGrd.Cols = 1
        tGrd.Cols = rs.Fields.count + 1
        For i = 1 To tGrd.Cols - 1
            tGrd.ColKey(i) = rs.Fields(i - 1).Name
            tGrd.TextMatrix(0, i) = rs.Fields(i - 1).Name
        Next
    End If

    tGrd.Redraw = flexRDNone
    i = tGrd.rows

    Do While Not rs.EOF
        tGrd.AddItem i - tGrd.FixedRows + 1
        For j = 0 To rs.Fields.count - 1
            If tGrd.ColIndex(rs.Fields(j).Name) <> -1 Then
                If rs.Fields(j).type = adCurrency Then
                    tGrd.TextMatrix(i, tGrd.ColIndex(rs.Fields(j).Name)) = val(rs.Fields(j).value & "")
                Else
                    tGrd.TextMatrix(i, tGrd.ColIndex(rs.Fields(j).Name)) = Trim(rs.Fields(j).value & "")
                End If
            End If
        Next
        i = i + 1
        rs.MoveNext
    Loop

    If InsertRow Then tGrd.AddItem tGrd.rows - tGrd.FixedRows + 1

    tGrd.Redraw = flexRDDirect
End Sub

 Public Sub saveGrid(ByVal Sqlstmt As String, ByRef tGrd As Object, ByVal ChekPoint As String, ByVal Index As String, ParamArray FieldValue())
    On Error GoTo Err
    Dim tRs As New ADODB.Recordset
    Dim s As String
    Dim mIndex As Long
    
    Dim mLastIndex As String
    If Index <> "" Then
        If mId(Index, 1, 5) = "Index" Then
            mLastIndex = mId(Index, 6)
            Index = "Id"
        End If
    End If
    
    
    tRs.Open Sqlstmt, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    ' *******************************************
    Dim ii As Long, i As Long
    If val(mLastIndex) = 0 Then
        ii = 0
    Else
        ii = val(mLastIndex)
    End If
    For i = tGrd.FixedRows To tGrd.rows - 1
        If ChekPoint <> "" Then
            If Trim(tGrd.TextMatrix(i, tGrd.ColIndex(ChekPoint))) = "" Then GoTo NextStep
        End If
        '**********************
        tRs.AddNew
        ii = ii + 1
        If Index <> "" And Index <> "IncrementID" Then tRs(Index) = ii
        Dim k As Long
        For k = 0 To UBound(FieldValue) Step 2
            tRs.Fields.Item(FieldValue(k)).value = FieldValue(k + 1)
            'Debug.Print FieldValue(k) & " " & tRs.Fields.Item(FieldValue(k)).Value
        Next
        '*************************
        'Debug.Print "fields count " & tRs.Fields.count
        Dim j As Long
        For j = 0 To tRs.Fields.count - 1

            If tGrd.ColIndex(tRs.Fields.Item(j).Name) <> -1 Then
            If tRs.Fields.Item(j).Name = "GroupUniqueFileMaster" Then
                j = j
            End If
            If Index = "IncrementID" And tRs.Fields.Item(j).Name = "RowId" Then
                                If tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = "" Then
                                    tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = GenerateGUID
                                    tRs.Fields.Item(j).value = "{" & tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) & "}"
                                Else
                                    If InStr(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)), "{") Then
                                    Else
                                        tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) = "{" & tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)) & "}"
                                        tRs.Fields.Item(j).value = tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))
                                    End If
                                End If
                            End If
                If tRs.Fields.Item(j).type = adInteger Or tRs.Fields.Item(j).type = adCurrency Or tRs.Fields.Item(j).type = adBoolean Or tRs.Fields.Item(j).type = adSmallInt Or tRs.Fields.Item(j).type = adBigInt Or tRs.Fields.Item(j).type = adTinyInt Or tRs.Fields.Item(j).type = adUnsignedTinyInt Or tRs.Fields.Item(j).type = adNumeric Or tRs.Fields.Item(j).type = adDouble Or tRs.Fields.Item(j).type = adDecimal Then
                    If tRs.Fields.Item(j).type = adBoolean Then
                        tRs.Fields.Item(j).value = (UCase(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "TRUE") Or (UCase(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = "-1") Or (val(tGrd.ValueMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) = -1)
                    Else
'                        If tGrd.ColComboList(tGrd.ColIndex(tRS.Fields.Item(j).Name)) <> "" Then
'                            tRS.Fields.Item(j).Value = tGrd.ValueMatrix(i, tGrd.ColIndex(tRS.Fields.Item(j).Name))
'                        Else
                            'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                            On Error Resume Next
                            
                            If Index = "IncrementID" And tRs.Fields.Item(j).Name = "ID" Then
                                Index = "IncrementID"
                                
                            Else
                                tRs.Fields.Item(j).value = val(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name)))
                            End If
                            
                            'End If
'                        End If
                    End If
                Else
                    If tRs.Fields.Item(j).type = adDBTimeStamp Or tRs.Fields.Item(j).type = adDBTime Or tRs.Fields.Item(j).type = adDBDate Then
                        If Not IsDate(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))) Then
                            tRs.Fields.Item(j).value = Null
                        Else
                            tRs.Fields.Item(j).value = tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name))
                        End If
              
              
              ElseIf tRs.Fields.Item(j).type = adGUID Then
    Dim strGuidValue As String
    strGuidValue = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))

    ' Check if braces are missing and add them
    If Len(strGuidValue) > 0 Then
        If left(strGuidValue, 1) <> "{" Then strGuidValue = "{" & strGuidValue
        If right(strGuidValue, 1) <> "}" Then strGuidValue = strGuidValue & "}"

        On Error Resume Next
        tRs.Fields.Item(j).value = strGuidValue
        If Err.Number <> 0 Then
            Debug.Print "Error assigning GUID: " & Err.Description & " for value: " & strGuidValue
            Err.Clear
            ' Consider setting to Null if conversion fails, or raise a more specific error
            tRs.Fields.Item(j).value = Null
        End If
        'On Error GoTo Err
    Else
        tRs.Fields.Item(j).value = Null
    End If
'End If

'                        tRs.Fields.Item(j).value = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))
                    Else
                        'If Index <> "" And UCase(tRs.Fields.Item(j).Name) <> UCase(tRs(Index).Name) Then
                        tRs.Fields.Item(j).value = Trim(tGrd.TextMatrix(i, tGrd.ColIndex(tRs.Fields.Item(j).Name) & ""))
                        'End If
                    End If
                End If
            End If
            'Debug.Print tRs.Fields.Item(j).Name & " = " & tRs.Fields.Item(j).Value
        Next
tRs.update
NextStep:
    Next
    tRs.Close
    Exit Sub
Err:
    If Err.Number = -2147217887 Then        ' one item is empty
        Resume Next
    End If
    '    Resume Next
End Sub



 





 
 
Public Function NormalizeArabic(ByVal Txt As String) As String
    Dim s As String
    s = Txt
    
    '  ÊÕÌœ «·√·›« 
    s = Replace(s, "√", "«")
    s = Replace(s, "≈", "«")
    s = Replace(s, "¬", "«")
    
    '  ÊÕÌœ «·Ì«¡ Ê«·√·› «·„ﬁ’Ê—…
    s = Replace(s, "Ï", "Ì")
    s = Replace(s, "∆", "Ì")
    
    '  ÊÕÌœ «· «¡ «·„—»Êÿ…
    s = Replace(s, "…", "Â") ' √Ê Œ·ÌÂ« " " Õ”» „« Ì‰«”»ﬂ
    
    '  ÊÕÌœ «·Ê«Ê Ê«·Â„“…
    s = Replace(s, "ƒ", "Ê")
    
    ' ≈“«·… «· ÿÊÌ·
    s = Replace(s, "‹", "")
    
    ' ≈“«·… «· ‘ﬂÌ· (Õ—ﬂ« )
    s = Replace(s, "Û", "")
    s = Replace(s, "", "")
    s = Replace(s, "ı", "")
    s = Replace(s, "Ò", "")
    s = Replace(s, "ˆ", "")
    s = Replace(s, "Ú", "")
    s = Replace(s, "˙", "")
    s = Replace(s, "¯", "")
    
    NormalizeArabic = s
End Function

