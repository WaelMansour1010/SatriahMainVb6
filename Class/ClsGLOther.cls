VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGLOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sat
Public FlgVat As Integer
Public Vatyo As Double
Public Vat As Double
Public AccountCode2 As String
Public CurrRow As Integer
Public UnitString As String

Public TotalValue As Double
Public SupplierID As Integer
Public CusVATNO As String
Public SupplierName As String
Public PriceTotal As Double
Public lineTotal As Double
Public Rate As Double
Public DescAccount As String
Public NextAccount_Code As String
Public BillNo As String
'Public NetCommession As Double
'Public NetCommessionVAt As Double
Public uuidCounter As Double
Public Id700 As String

Dim discountvalue As Double
     Dim PriceAmount As Double
     Dim InvoiceQuantity As Double
     Dim Totallinewithvat As Double
          Dim Totallinewithoutvat As Double
                   Dim TotalVat As Double
'Public NetCommession As Double
'Public NetCommessionVAt As Double
Public Transaction_ID As Double
Public ID As String
Public Invoicetype As Integer

Public ActivityTypeIdInvoice As Integer
Public branch_idInvoice As Integer

Public fromscreen As Integer



Public Unitss As String
Public StrUnit As String
Public uintid As Integer
Public mType As Integer
Public iqarid As Integer

Public IssueDate As String
Public chkTaxExempt As Integer
Public order_no As String

Public IssueTim As String
Public InvoiceTypeCodeID As String
Public InvoiceTypeCodename As String
Public DocumentCurrencyCode As String
Public TaxCurrencyCode As String
Public InvoiceDocumentReferenceID As String
Public AdditionalDocumentReferenceICVUUID As String
Public ActualDeliveryDate As String
Public LatestDeliveryDate As String
Public PaymentMeansCode As String
Public InstructionNote As String
Public PayeeFinancialAccount As String
Public paymentnote As String
Public Identificationid As String
Public schemeID As String
Public StreetName As String
Public AdditionalStreetName As String
Public BuildingNumber As String
Public PlotIdentification As String
Public CityName As String
Public PostalZone As String
Public CountrySubentity As String
Public CitySubdivisionName As String
Public IdentificationCode As String
Public RegistrationName As String
Public CompanyID As String
Public allowancechargeAmount As Double
Public AllowanceChargeReason As String
Public TaxCategoryID As String
Public TaxCategoryPercent As Double
Public PayableAmount As Double
'Public Totalvat As Double
Public PrepaidAmount As Double
Public InvoiceHash
Public SingedXML
Public EncodedInvoice
Public UUID
Public isExpens As Boolean
Public QRCode
Public PIH
Public Export
Public SingedXMLFileName
Public errorMessage
Public zatcaStatus
Public QrCodeDataPath
Public warrningmessage
Dim str As String
Public ErrorMessageS
Public docType As String
Public docType2 As String
Dim usingformat As Boolean
Public Simplevalue As Boolean
Public SendingMode As Integer

 Public POSConnection  As ADODB.Connection
Private cn2 As ADODB.Connection
    Function savewarrninganderror(Transaction_ID As Double, Optional warrningmessage As String, Optional errorMessage As String, Optional ByVal mTableName As String = "transactions", Optional ByVal mFieldIDName As String = "Transaction_ID")
    Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    If mTableName = "tblEInvoice" Then
        'StrSQL = "SELECT  *  from " & mTableName & "  Where (" & mFieldIDName & "  = '" & Transaction_ID & "')"
        StrSQL = "SELECT  *  from " & mTableName & "  Where  ID  = " & Transaction_ID
    Else
        StrSQL = "SELECT  *  from " & mTableName & "  Where (" & mFieldIDName & "  = " & Transaction_ID & ")"
    End If
    RsDevsub.Open StrSQL, cn2, adOpenKeyset, adLockOptimistic, adCmdText
    If RsDevsub.RecordCount > 0 Then
    If warrningmessage <> "" And Not IsEmpty(warrningmessage) Then
     RsDevsub("warrningmessage").value = warrningmessage
     End If
     If errorMessage <> "" And Not IsEmpty(ErrorMessageS) Then
     RsDevsub("ErrorMessageS").value = errorMessage
     End If
     
   RsDevsub.update
   End If
    RsDevsub.Close
    End Function
Function updatedatatable(Optional ByVal docType As Integer = 0, Optional ByVal mTableName As String = "transactions", Optional ByVal mFieldIDName As String = "Transaction_ID")

str = " update transactionsVatDetails  set isdeleted  =1 where Transaction_ID =" & Transaction_ID & " and docType = " & docType
cn2.Execute str


 
Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from transactionsVatDetails Where (1 = -1)"
    RsDevsub.Open StrSQL, cn2, adOpenKeyset, adLockOptimistic, adCmdText
 RsDevsub.AddNew
 RsDevsub("InvoiceHash").value = InvoiceHash
 RsDevsub("SingedXML").value = SingedXML
 RsDevsub("EncodedInvoice").value = EncodedInvoice
 RsDevsub("UUID").value = UUID
 RsDevsub("QRCode").value = QRCode
 RsDevsub("docType").value = docType
  RsDevsub("TableName").value = mTableName
 
 RsDevsub("PIH").value = PIH
 RsDevsub("SingedXMLFileName").value = SingedXMLFileName
 RsDevsub("Transaction_ID").value = Transaction_ID
RsDevsub("uuidCounter").value = uuidCounter
 RsDevsub.update
 
 'Cn2.Execute " update  transactions  set QrCodeImage=  (select QrCodeImage from transactionsVatDetails whereTransaction_ID=" & Transaction_ID & "   ) where Transaction_ID= " & Transaction_ID
 Dim Picture1 As PictureBox
 If mTableName = "tblEInvoice" Then
        SaveQRCode2 "transactionsVatDetails", "Transaction_ID", CLng(Transaction_ID), order_no, IssueDate, "", Picture1, 0, , , CStr(QRCode), val(docType), fromscreen
 Else
 SaveQRCode2 "transactionsVatDetails", "Transaction_ID", CLng(Transaction_ID), ID, IssueDate, "" _
         , Picture1, 0, , , CStr(QRCode), val(docType), fromscreen
End If
'   If mTableName = "tblEInvoice" Then
'
'        cn2.Execute " update  " & mTableName & "  set QrCodeImage=  (select QrCodeImage from transactionsVatDetails where    isnull(isdeleted,0)=0 and Transaction_ID =" & Transaction_ID & " and docType = " & docType & "  ) where " & mFieldIDName & " = '" & Transaction_ID & "'"
'    Else
        cn2.Execute " update  " & mTableName & "  set QrCodeImage=  (select QrCodeImage from transactionsVatDetails where    isnull(isdeleted,0)=0 and Transaction_ID =" & Transaction_ID & " and docType = " & docType & "  ) where " & mFieldIDName & " = " & Transaction_ID
    'End If
        
'    SaveQRCode2 "transactions", "Transaction_ID", CLng(Transaction_ID), ID, IssueDate, "" _
'         , Picture1, 0, , , CStr(QRCode), val(docType), fromscreen
End Function

Function CalllastPIH(Optional ByRef X As Double) As String
Dim lastpih As String
Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  top(1) InvoiceHash ,uuidCounter from transactionsVatDetails  Where not ( InvoiceHash is null)  and ISNULL(CAST(InvoiceHash AS NVARCHAR(MAX)), '') <> '' order by id desc"
    RsDevsub.Open StrSQL, cn2, adOpenKeyset, adLockOptimistic, adCmdText
    If RsDevsub.RecordCount > 0 Then
            lastpih = IIf(IsNull(RsDevsub("InvoiceHash").value), "", RsDevsub("InvoiceHash").value)
            uuidCounter = IIf(IsNull(RsDevsub("uuidCounter").value), 0, RsDevsub("uuidCounter").value)
    Else
            lastpih = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="
            uuidCounter = 0
    End If
 CalllastPIH = lastpih
 uuidCounter = uuidCounter + 1
  RsDevsub.Close
 
End Function


Function CalllastPIH2() As String
Dim lastpih As String
Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  top(1) InvoiceHash  from transactionsVatDetails  Where not ( pih is null) and ISNULL(CAST(InvoiceHash AS NVARCHAR(MAX)), '') <> '' order by id desc"
    
    RsDevsub.Open StrSQL, cn2, adOpenKeyset, adLockOptimistic, adCmdText
    If RsDevsub.RecordCount > 0 Then
            lastpih = IIf(IsNull(RsDevsub("InvoiceHash").value), "", RsDevsub("InvoiceHash").value)
    Else
            lastpih = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="
    End If
 CalllastPIH2 = lastpih
End Function

'
'
'
Function generateInvoice(Optional ByVal docType As Integer = 0, Optional ByVal mTableName As String = "transactions", Optional ByVal mFieldIDName As String = "Transaction_ID", Optional ByVal ActivityTypeIdInv As Integer = 0, Optional ByVal branch_idInv As Integer = 0)

 On Error Resume Next: Err.Clear
' Dim ubl As ZatcaUBLXMLGenerator.UBLXML
'    Set ubl = New ZatcaUBLXMLGenerator.UBLXML
'Dim inv As ZatcaUBLXMLGenerator.invoice
'    Set inv = New ZatcaUBLXMLGenerator.invoice
'    Dim res As ZatcaUBLXMLGenerator.Result
'    Set res = New ZatcaUBLXMLGenerator.Result
'    '
     Dim ubl As Object
      Dim SendingMode As Integer
      
      Dim mDiscountProject As Double
     Dim SecretKey As String
 Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
    Dim inv As Object
 Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
  Dim res As Object
 Set res = CreateObject("ZatcaUBLXMLGenerator.Result")
    
     If cschemeID = "" Then
    cschemeID = "CRN"
 End If
 
 
  If SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
        Set cn2 = POSConnection
        
    Else
        Set cn2 = Cn
       
    End If
'    »Ì«‰«  «·›« Ê—…
'    QRCode = ""

If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

'ID = 12345
    inv.ID = ID ' "1230"
    If IssueDate = "" Then IssueDate = Date
    inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")           '"2021-01-05" 'IssueDate
    If IssueTim = "" Then IssueTim = Now
    inv.IssueTime = Format(IssueTim, "HH:MM:SS")   '"09:32:40"  ' Format ( IssueTime,"HH:MM:SS")
     ' ›« Ê—… ÷—Ì»Ì… «Ê „»”ÿ… 388
      ' «‘⁄«— „œÌ‰ 383
      ' 381 «‘⁄«— œ«∆‰
    inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
  inv.InvoiceTypeCode.Name = InvoiceTypeCodename
    'inv.invoiceTypeCode.Name based on format NNPNESB
    'NN 01 ··›« Ê—… «·÷—Ì»Ì…
    'NN 02 ··›« Ê—… «·÷—Ì»Ì… «·„»”ÿ…
    'P ›Ï Õ«·… ›« Ê—… ·ÿ—› À«·À ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'N ›Ï Õ«·… ›« Ê—… «”„Ì… ‰ﬂ » 1 ›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'E ›Ï Õ«·… ›« Ê—… ··’«œ—«  ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'S ›Ï Õ«·… ›« Ê—… „·Œ’… ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'B  ›Ï Õ«·… ›« Ê—… –« Ì… ‰ﬂ » 1
    'B ›Ï Õ«·… «‰ «·›« Ê—… ’«œ—« =1 ·« Ì„ﬂ‰ «‰  ﬂÊ‰ «·›« Ê—… –« Ì… =1
If DocumentCurrencyCode = "" Then DocumentCurrencyCode = "SAR"
If TaxCurrencyCode = "SR" Or TaxCurrencyCode = "" Then TaxCurrencyCode = "SAR"
If DocumentCurrencyCode = "SR" Or DocumentCurrencyCode = "" Then DocumentCurrencyCode = "SAR"

    inv.DocumentCurrencyCode = DocumentCurrencyCode
    inv.TaxCurrencyCode = TaxCurrencyCode
   'inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="
       'Ì „ «” œ⁄« «Œ— PIH
   'CalllastPIH
   inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
 
 
    '›Ï Õ«·… «‘⁄«— œ«∆‰ «Ê „œÌ‰ ›ﬁÿ Â«‰ﬂ » —ﬁ„ «·›« Ê—… «··Ï «’œ—‰« «·«‘⁄«— ·ÌÂ«
 'inv.BillingReference.InvoiceDocumentReferenceID = val(InvoiceDocumentReferenceID)
 
    
 
    Dim InvoiceDocumentReference As Object
    Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
    InvoiceDocumentReference.ID = val(InvoiceDocumentReferenceID)
    inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    ' "?Invoice Number: 354; Invoice Issue Date: 2021-02-10?"
    ' ﬁÌ„… ⁄œ«œ «·›« Ê—…
    inv.AdditionalDocumentReferenceICV.UUID = uuidCounter ' 12345 val(AdditionalDocumentReferenceICVUUID) '„„ﬂ‰ ÌﬂÊ‰ —ﬁ„ «·›« Ê—… Ê·ﬂ‰ ·«»œ ÌﬂÊ‰ «—ﬁ«„ ›ﬁÿ
    ' ›Ï Õ«·… ›« Ê—… „»”ÿ… Ê›« Ê—… „·Œ’… Â«‰ﬂ »  «—ÌŒ «· ”·Ì„ Ê«Œ—  «—ÌŒ «· ”·Ì„

If ActualDeliveryDate = "" Then ActualDeliveryDate = IssueDate
If LatestDeliveryDate = "" Then LatestDeliveryDate = IssueDate
    inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")   ' ActualDeliveryDate ' "2022-10-22"
  If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")   ' LatestDeliveryDate '"2022-10-23"
    End If
   ' »Ì«‰«  «·œ›⁄
   ' «ﬂÊ«œ „⁄Ì‰…
   ' «Œ Ì«—Ï ﬂÊœ «·œ›⁄
       Dim paymentmean As Object
     Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
    paymentmean.PaymentMeansCode = PaymentMeansCode ' ????? ????? ????
     
    If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then '«‘⁄«— „œÌ‰ «Ê œ«∆‰
   paymentmean.InstructionNote = "Returned items" '›Ï Õ«·… «‘⁄«— «·œ«∆‰ Ê«·„œÌ‰ ·«»œ „‰ «œŒ«·Â«
   End If
      
    paymentmean.PayeeFinancialAccount.ID = PayeeFinancialAccount ' ?? ???? ??? ????? ?? ???? ????? ??? ???? ????? IBAN
    paymentmean.PayeeFinancialAccount.paymentnote = paymentnote ' ?? ???? ??? ????? ?? ???? ?????
    inv.PaymentMeans.Add paymentmean
    
     

 

 
    inv.SupplierParty.PartyIdentification.ID = cIdentificationID
    If Identificationid = "" Then
        'cschemeID ="OTH"
    End If
    
    inv.SupplierParty.PartyIdentification.schemeID = cschemeID
    inv.SupplierParty.PostalAddress.StreetName = cStreetName
    inv.SupplierParty.PostalAddress.AdditionalStreetName = cAdditionalStreetName ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.BuildingNumber = cBuildingNumber
    inv.SupplierParty.PostalAddress.PlotIdentification = cPlotIdentification ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.CityName = cCityName
    inv.SupplierParty.PostalAddress.PostalZone = cPostalZone
    inv.SupplierParty.PostalAddress.CountrySubentity = cCountrySubentity ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.CitySubdivisionName = cCitySubdivisionName
    inv.SupplierParty.PostalAddress.Country.IdentificationCode = "SA" 'cIdentificationCode
    inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim(Companynameforvat)
    inv.SupplierParty.PartyTaxScheme.CompanyID = cCompanyID

If mId(InvoiceTypeCodename, 1, 2) = "01" Then
    '»Ì«‰«  «·⁄„Ì·
    If Id700 <> "" Then
        inv.CustomerParty.PartyIdentification.ID = Id700  'Company_Comment
    Else
        
       ' If Identificationid = "" Then Identificationid = "N/A"
        
           
    
        inv.CustomerParty.PartyIdentification.ID = Identificationid 'Company_Comment
    End If
    
     If CompanyID = "" Then
                CompanyID = "300000000000003"
                schemeID = "OTH"
            End If
    
    inv.CustomerParty.PartyIdentification.schemeID = schemeID '
    If Identificationid = "N/A" Or Identificationid = "" Then
        
        If StreetName = "" Then
            StreetName = "N/A"
        End If
        If AdditionalStreetName = "" Then
            AdditionalStreetName = "N/A"
        End If
        If BuildingNumber = "" Then
            BuildingNumber = "N/A"
        End If
        If CityName = "" Then
            CityName = "N/A"
        End If
        If PostalZone = "" Then
            PostalZone = "00000"
        End If
        If CountrySubentity = "" Then
            CountrySubentity = "N/A"
        End If
        If CitySubdivisionName = "" Then
            CitySubdivisionName = "N/A"
        End If
    End If
    If IdentificationCode = "" Then IdentificationCode = "SA"
    
    inv.CustomerParty.PostalAddress.StreetName = StreetName
    inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName ' «Œ Ì«—Ì
    inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
    inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification ' «Œ Ì«—Ì
    
    inv.CustomerParty.PostalAddress.CityName = CityName
    inv.CustomerParty.PostalAddress.PostalZone = PostalZone
    inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity '«”„ «·Ê·«Ì… «Œ Ì«—Ï
    inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim(RegistrationName) 'Company_Name_Eng
    If Id700 <> "" Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = "N/A" 'VATRegNo
        inv.CustomerParty.PartyTaxScheme.CompanyID = "" 'VATRegNo
    Else
        'If CompanyID = "" Then CompanyID = "N/A"
        inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID 'VATRegNo
        
    End If
End If




If mId(InvoiceTypeCodename, 1, 2) = "02" And (mTableName = "tblEInvoice" Or mTableName = "PropertyDueBatchDetail" Or mTableName = "DebitAndCreditNotification") Then
    '»Ì«‰«  «·⁄„Ì·
        If CompanyID = "" Then
                CompanyID = "300000000000003"
                schemeID = "OTH"
            End If
    If Identificationid = "0" Then Identificationid = ""
    If Id700 <> "" Then
        inv.CustomerParty.PartyIdentification.ID = Id700  'Company_Comment
    Else
        inv.CustomerParty.PartyIdentification.ID = Identificationid 'Company_Comment
    End If
    inv.CustomerParty.PartyIdentification.schemeID = schemeID '
    inv.CustomerParty.PostalAddress.StreetName = StreetName
    inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName ' «Œ Ì«—Ì
    inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
    inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification ' «Œ Ì«—Ì
    
    inv.CustomerParty.PostalAddress.CityName = CityName
    inv.CustomerParty.PostalAddress.PostalZone = PostalZone
    inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity '«”„ «·Ê·«Ì… «Œ Ì«—Ï
    inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
    If IdentificationCode = "" Then IdentificationCode = "SA"
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
    inv.CustomerParty.PartyLegalEntity.RegistrationName = RegistrationName 'Company_Name_Eng
    If Id700 <> "" Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = "N/A" 'VATRegNo
        inv.CustomerParty.PartyTaxScheme.CompanyID = "" 'VATRegNo
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID 'VATRegNo
    End If
End If


    ' Â‰« ‰ﬁœ— ‰⁄„· loop Ê‰·› ⁄·Ï «·ﬂÊœ ÊÌ√„« ‰⁄„·Â Œ’„ «Ê «÷«›…
    ' »Ì«‰«  «·Œ’„ ⁄·Ï „” ÊÏ «·›« Ê—…
    Dim allowancecharge As Object
    ' ﬂÊœ «·Œ’„
    Set allowancecharge = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'allowancechargeAmount = 0 '   0800000
 
    allowancecharge.Amount = allowancechargeAmount ' Format(allowancechargeAmount, ".##")     ' ﬁÌ„… «·Œ’„
  ' allowancecharge.MultiplierFactorNumeric = 10
  ' allowancecharge.BaseAmount = 100

    allowancecharge.AllowanceChargeReason = "Discount" ' ”»» «·Œ’„ ⁄·Ï „” ÊÏ «·›« Ê—…
    If chkTaxExempt = 1 Or TaxCategoryPercent = 0 Then
        allowancecharge.taxCategory.ID = "E" ' ﬂÊœ «·÷—Ì»…
        allowancecharge.taxCategory.Percent = 0# ' ‰”»… «·÷—Ì»…
    Else
    
        allowancecharge.taxCategory.ID = "S" ' ﬂÊœ «·÷—Ì»…
        allowancecharge.taxCategory.Percent = 15# ' ‰”»… «·÷—Ì»…
    End If
    inv.allowanceCharges.Add allowancecharge

    ' «Ã„«·Ï «·„»·€ »⁄œ «·÷—Ì»… - «·„»·€ «·„œ›Ê⁄
   inv.LegalMonetaryTotal.PayableAmount = (PayableAmount)
  ' «·„»·€ «·„œ›Ê⁄
     If val(InvoiceTypeCodeID) = 388 Then
   inv.LegalMonetaryTotal.PrepaidAmount = (PrepaidAmount)
   End If
   '********************************************************************************

   '**************************************************************************
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    

     Dim rs As New ADODB.Recordset
     Dim StrSQL As String
 If docType = 0 Then
 
 Dim Percetage As Double

StrSQL = "SELECT dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, TblItems.ItemWithOutVAT, "
StrSQL = StrSQL & "dbo.Transaction_Details.ShowQty, PerforValue = 0,  "
StrSQL = StrSQL & "dbo.Transaction_Details.ItemDiscount AS discountvalue, Transactions.Trans_Discount as discountvalue2,"

' «· Õﬁﬁ „„« ≈–« ﬂ«‰ «·”⁄— Ì‘„· «·÷—Ì»…
If SystemOptions.PriceWithVAT Then
    PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage
    
    ' Õ”«» «·”⁄— »œÊ‰ «·÷—Ì»… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100)) "
    StrSQL = StrSQL & "ELSE dbo.Transaction_Details.showPrice "
    StrSQL = StrSQL & "END, 2) AS showPrice, "

    ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent, "

    ' ‰”»… «·÷—Ì»… «·„√ŒÊ–… „‰ «·„ €Ì— Ê·Ì” „‰ «·ÃœÊ·
    'StrSQL = StrSQL & " CAST(" & Percetage & " AS DECIMAL(18, 2)) AS vatPercent, "

    '  ÕœÌœ „« ≈–« ﬂ«‰ «·„‰ Ã Œ«÷⁄« ··÷—Ì»… √„ ·«
    StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN " & Percetage
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "
Else
    ' ›Ì Õ«·… «·”⁄— €Ì— ‘«„· «·÷—Ì»…
    StrSQL = StrSQL & "dbo.Transaction_Details.showPrice , "
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100), 2) AS VATAmount, "
    StrSQL = StrSQL & " ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent, dbo.Transaction_Details.Vatyo "
End If

'  ﬂ„·… «·«” ⁄·«„ »—»ÿ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.Transactions "
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
StrSQL = StrSQL & " WHERE (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"




'Dim Percetage As Double



StrSQL = "SELECT dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, TblItems.ItemWithOutVAT, "
StrSQL = StrSQL & "dbo.Transaction_Details.ShowQty, PerforValue = 0,  "
StrSQL = StrSQL & "dbo.Transaction_Details.ItemDiscount + dbo.Transaction_Details.TotalDiscountPerLine AS discountvalue, "
StrSQL = StrSQL & "Transactions.Trans_Discount AS discountvalue2, Transactions.Trans_DiscountType, "

' «· Õﬁﬁ „„« ≈–« ﬂ«‰ «·”⁄— Ì‘„· «·÷—Ì»…
If SystemOptions.PriceWithVAT Then
    PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage

    ' Õ”«» ≈Ã„«·Ì «·”⁄— ﬁ»· «·Œ’„
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty, 2) AS totalPrice, "

    ' Õ”«» ≈Ã„«·Ì «·”⁄— »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0), 2) AS totalPriceAfterDiscount, "

    ' «” Œ—«Ã «·”⁄— »œÊ‰ «·÷—Ì»… „‰ «·≈Ã„«·Ì »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "/ (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100)), 2) / dbo.Transaction_Details.ShowQty AS showPrice, "

    ' Õ”«» «·÷—Ì»… »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "- (((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "/ (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))), 2) AS vatPercent, "
    
    
      ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN IsNull(TblItems.ItemWithOutVAT,0) = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent2, "
    
      StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN  IsNull(TblItems.ItemWithOutVAT,0) = 0 THEN " & Percetage
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "


Else
    ' ≈–« ﬂ«‰ «·”⁄— €Ì— ‘«„· «·÷—Ì»…
   ' StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice, 2) AS showPrice, "


 StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0))  , 2) / dbo.Transaction_Details.ShowQty AS showPrice,"
    
    
    ' Õ”«» ≈Ã„«·Ì «·”⁄— ﬁ»· «·Œ’„
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty, 2) AS totalPrice, "

    ' Õ”«» ≈Ã„«·Ì «·”⁄— »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0), 2) AS totalPriceAfterDiscount, "

    ' Õ”«» «·÷—Ì»… »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "* (CAST(dbo.Transaction_Details.Vatyo AS DECIMAL(18, 2)) / 100), 2) AS vatPercent, "
    
    
      ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN IsNull(ItemWithOutVAT,0) = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST( dbo.Transaction_Details.Vatyo AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent2, "
    
      StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN IsNull(ItemWithOutVAT,0) = 0 THEN Vatyo"
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "

End If

' ‰”»… «·÷—Ì»…


'  ÕœÌœ „« ≈–« ﬂ«‰ «·„‰ Ã Œ«÷⁄« ··÷—Ì»… √„ ·«


'  ﬂ„·… «·«” ⁄·«„ »—»ÿ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.Transactions "
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
StrSQL = StrSQL & " WHERE (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"



'    StrSQL = "SELECT     dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,TblItems.ItemWithOutVAT, dbo.Transaction_Details.ShowQty,PerforValue = 0, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemDiscount as discountvalue, "
'      If SystemOptions.PriceWithVAT Then
'        PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage
'        StrSQL = StrSQL & "                     ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent,   "
'
'        StrSQL = StrSQL & "  CASE"
'        StrSQL = StrSQL & "  WHEN TblItems.ItemWithOutVAT = 0 THEN " & Percetage
'        StrSQL = StrSQL & "  ELSE 0"
'        StrSQL = StrSQL & "  END As Vatyo"
'
'
'    Else
'        StrSQL = StrSQL & "                     ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent, dbo.Transaction_Details.Vatyo"
'    End If
'
'    StrSQL = StrSQL & "    FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "                          dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    StrSQL = StrSQL & "                          dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
'    StrSQL = StrSQL & "     Where (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"

allowancecharge.Amount = 0
ElseIf docType = 1 Then

    

     StrSQL = " SELECT     project_bill_details.item as   ItemName, project_bill_details.item as ItemNamee, ItemWithOutVAT = 0, 1 as ShowQty,project_billl.PerforValue,"
     
     
     
   StrSQL = StrSQL & "               showPrice = IsNull(dbo.project_bill_details.totEx,0) - (IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0))"
   StrSQL = StrSQL & "               ,(IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0)) AS discountvalue"
   StrSQL = StrSQL & "               ,ROUND(IsNull(dbo.project_bill_details.LineVat,0), 2) AS vatPercent22"
   StrSQL = StrSQL & "               ,15 AS Vatyo"
   StrSQL = StrSQL & "               ,ROUND(( dbo.project_bill_details.totEx - (IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0)) ) * .15, 2) AS vatPercent"
     
     'StrSQL = StrSQL & "                         showPrice  = dbo.project_bill_details.totEx  - (IsNull(project_billl.discount,0) + isnull(project_billl.DiscountGMater,0) + isnull(project_billl.advancedPayment,0)) , "
'     StrSQL = StrSQL & "                         showPrice  = dbo.project_bill_details.totEx  ROUND((dbo.project_bill_details.totEx - (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment))   , "
'     StrSQL = StrSQL & "                         (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment) as discountvalue,"
'     StrSQL = StrSQL & "                         ROUND(dbo.project_bill_details.LineVat, 2) AS vatPercent22, 15 as Vatyo,"
'    ' StrSQL = StrSQL & "                         ROUND((dbo.project_bill_details.totEx -( project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment)) * .15,2) AS vatPercent"
'     StrSQL = StrSQL & "                         ROUND((dbo.project_bill_details.totEx ) * .15,2) AS vatPercent,"
'
'     StrSQL = StrSQL & "                    ROUND((dbo.project_bill_details.totEx - (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment)) * 0.15, 2) As VatPercent"
     StrSQL = StrSQL & "             FROM         dbo.project_billl INNER JOIN"
     StrSQL = StrSQL & "     dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id"
     
     StrSQL = StrSQL & "     Where (dbo.project_billl.ID = " & Transaction_ID & ")"
     
     


' »‰«¡ «·«” ⁄·«„ «·√”«”Ì


' »‰«¡ «·«” ⁄·«„ «·√”«”Ì
StrSQL = " SELECT project_bill_details.item AS ItemName, "
StrSQL = StrSQL & " project_bill_details.item AS ItemNamee, "
StrSQL = StrSQL & " 0 AS ItemWithOutVAT, "
StrSQL = StrSQL & " 1 AS ShowQty, "
StrSQL = StrSQL & " project_billl.PerforValue, "

'  ⁄œÌ· Õﬁ· «·”⁄— »‰«¡ ⁄·Ï ﬁÌ„… SystemOptions.SuppCreat4Acc
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ ÿ—Õ «·Œ’„ «·„Ê“⁄
    'project_billl.Discount

   

    
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " CASE "
    StrSQL = StrSQL & " WHEN project_billl.discount1ID = 1 THEN "
    StrSQL = StrSQL & " dbo.project_bill_details.totEx * (IsNull(project_billl.discount1value, 0) / 100) "
    StrSQL = StrSQL & " ELSE "
    StrSQL = StrSQL & " IsNull(project_billl.discount1value, 0) "
    StrSQL = StrSQL & " END, "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ «” Œœ«„ «·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

' Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·”⁄— »⁄œ «·Œ’„
StrSQL = StrSQL & " vatPercent = ROUND((dbo.project_bill_details.totEx - "
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·Œ’„ «·„Ê“⁄
    StrSQL = StrSQL & " CASE "
    StrSQL = StrSQL & " WHEN project_billl.discount1ID = 1 THEN "
    StrSQL = StrSQL & " dbo.project_bill_details.totEx * (IsNull(project_billl.discount1value, 0) / 100) "
    StrSQL = StrSQL & " ELSE "
    StrSQL = StrSQL & " IsNull(project_billl.discount1value, 0) "
    StrSQL = StrSQL & " END "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ Õ”«» «·÷—Ì»… »«·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)) "
End If
StrSQL = StrSQL & ") * 0.15, 2), "

' ≈÷«›… »«ﬁÌ «·ÕﬁÊ·
StrSQL = StrSQL & " discountvalue = (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
StrSQL = StrSQL & " ROUND(IsNull(dbo.project_bill_details.LineVat, 0), 2) AS vatPercent22, "
StrSQL = StrSQL & " 15 AS Vatyo "

' «·—»ÿ »Ì‰ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.project_billl INNER JOIN "
StrSQL = StrSQL & " dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id "

' «·‘—ÿ
StrSQL = StrSQL & " WHERE (dbo.project_billl.ID = " & Transaction_ID & ")"


 
 ' »‰«¡ «·«” ⁄·«„ «·√”«”Ì
StrSQL = " SELECT project_bill_details.item AS ItemName, "
StrSQL = StrSQL & " project_bill_details.item AS ItemNamee, "
StrSQL = StrSQL & " 0 AS ItemWithOutVAT, "
StrSQL = StrSQL & " 1 AS ShowQty, "
StrSQL = StrSQL & " project_billl.PerforValue, "

'  ⁄œÌ· Õﬁ· «·”⁄— »‰«¡ ⁄·Ï ﬁÌ„… SystemOptions.SuppCreat4Acc
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „  Ê“Ì⁄ «·Œ’„ „‰ ›Ì·œ Discount ›ﬁÿ
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.Discount, 0) * "
    StrSQL = StrSQL & " (dbo.project_bill_details.totEx / "
    StrSQL = StrSQL & " (SELECT SUM(dbo.project_bill_details.totEx) "
    StrSQL = StrSQL & " FROM dbo.project_bill_details "
    StrSQL = StrSQL & " WHERE dbo.project_bill_details.bill_id = " & Transaction_ID & "))), "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ «” Œœ«„ «·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

' Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·”⁄— »⁄œ «·Œ’„
StrSQL = StrSQL & " vatPercent = ROUND((dbo.project_bill_details.totEx - "
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·Œ’„ «·„Ê“⁄
    StrSQL = StrSQL & " (IsNull(project_billl.Discount, 0) * "
    StrSQL = StrSQL & " (dbo.project_bill_details.totEx / "
    StrSQL = StrSQL & " (SELECT SUM(dbo.project_bill_details.totEx) "
    StrSQL = StrSQL & " FROM dbo.project_bill_details "
    StrSQL = StrSQL & " WHERE dbo.project_bill_details.bill_id = " & Transaction_ID & "))) "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ Õ”«» «·÷—Ì»… »«·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)) "
End If
StrSQL = StrSQL & ") * 0.15, 2), "

' ≈÷«›… »«ﬁÌ «·ÕﬁÊ·
If SystemOptions.SuppCreat4Acc = False Then
    StrSQL = StrSQL & " discountvalue = IsNull(project_billl.discount, 0), "
Else
    StrSQL = StrSQL & " discountvalue = (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

StrSQL = StrSQL & " ROUND(IsNull(dbo.project_bill_details.LineVat, 0), 2) AS vatPercent22, "
StrSQL = StrSQL & " 15 AS Vatyo "

' «·—»ÿ »Ì‰ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.project_billl INNER JOIN "
StrSQL = StrSQL & " dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id "

' «·‘—ÿ
StrSQL = StrSQL & " WHERE (dbo.project_billl.ID = " & Transaction_ID & ")"


ElseIf docType = 3 Then

    
' ===== Pick connection by source =====
    If LCase$(mTableName) = "debitandcreditnotification" And SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
        
        Set cn2 = POSConnection
          Dim pct As Double
        ' ????? ?????? ?? ?????? ?????? ???? SQL (??????? ?Vatyo)
        ' ?? VATPercentage ????? ?? 0.15? ??????? ??? 15 ???? SQL.

        StrSQL = ""
        StrSQL = StrSQL & "SELECT TOP 1 " & vbCrLf
        StrSQL = StrSQL & "  ItemName       = N'Credit/Debit Note', " & vbCrLf
        StrSQL = StrSQL & "  ItemNamee      = N'Credit/Debit Note', " & vbCrLf
        StrSQL = StrSQL & "  ItemWithOutVAT = CASE WHEN ISNULL(d.VATValue,0)=0 THEN 1 ELSE 0 END, " & vbCrLf
        StrSQL = StrSQL & "  ShowQty        = 1, " & vbCrLf
        StrSQL = StrSQL & "  PerforValue    = 0, " & vbCrLf
        StrSQL = StrSQL & "  showPrice      = ROUND(ISNULL(d.MoneyAmount,0), 2), " & vbCrLf
        StrSQL = StrSQL & "  discountvalue  = 0, " & vbCrLf
        StrSQL = StrSQL & "  vatPercent     = ROUND(ISNULL(d.VATValue,0), 2), " & vbCrLf
        StrSQL = StrSQL & "  Vatyo          = CASE " & vbCrLf
        StrSQL = StrSQL & "                      WHEN ISNULL(d.VATPercentage,0) > 1 THEN d.VATPercentage " & vbCrLf
        StrSQL = StrSQL & "                      ELSE d.VATPercentage * 100 " & vbCrLf
        StrSQL = StrSQL & "                    END " & vbCrLf
        StrSQL = StrSQL & "FROM dbo.DebitAndCreditNotification d " & vbCrLf
        StrSQL = StrSQL & "WHERE d.Id = " & Transaction_ID & vbCrLf

    Else
    
        StrSQL = " SELECT    'Notes'   ItemName, 'Notes' as ItemNamee, ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, Notes.Note_Value as showPrice, 0 as discountvalue,"
        StrSQL = StrSQL & "                         ROUND(dbo.Notes.VAT, 2) AS vatPercent, 15 as Vatyo"
        StrSQL = StrSQL & "             FROM         "
        StrSQL = StrSQL & "     dbo.Notes "
        
        StrSQL = StrSQL & "     Where (dbo.Notes.NoteID = " & Transaction_ID & ")"
        Set cn2 = Cn
    End If

    

ElseIf docType = 4 Then
     StrSQL = " SELECT   top 1  ExpensesRemark   ItemName, ExpensesRemark as ItemNamee,ItemWithOutVAT = 0, Count as ShowQty,0 as PerforValue, Notes.price as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.Notes.Note_Value  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.notes"
      StrSQL = StrSQL & " Where IsNull(notes.CusID, 0) = 0 And notes_all = " & Transaction_ID & ""
      
      
ElseIf docType = 10 Then

    
    StrSQL = "SELECT     ItemName,ItemWithOutVAT = 0, Qty ShowQty,PerforValue = 0, Price showPrice, discountvalue = 0, "
    StrSQL = StrSQL & "                      VatValue as Vatyo ,"
    StrSQL = StrSQL & "                vatPercent =   IsNull(VatValue,0) / (Price *Qty ) * 100  , "
    StrSQL = StrSQL & "                Vatyo =   IsNull(VatValue,0) / (Price *Qty ) * 100   "
    
    StrSQL = StrSQL & "    FROM         tblEInvoice2"
    StrSQL = StrSQL & "     Where (dbo.tblEInvoice2.InvoiceID = '" & order_no & "')"
    
    
        StrSQL = ""
    StrSQL = StrSQL & "SELECT "
    StrSQL = StrSQL & "    ItemName, "
    StrSQL = StrSQL & "    ItemWithOutVAT = 0, "
    StrSQL = StrSQL & "    Qty AS ShowQty, "
    StrSQL = StrSQL & "    PerforValue = 0, "
    StrSQL = StrSQL & "    Price AS showPrice, "
    StrSQL = StrSQL & "    discountvalue = 0, "
    StrSQL = StrSQL & "    CASE "
    StrSQL = StrSQL & "        WHEN ISNULL(VatValue,0) = 0 THEN 0 "
    StrSQL = StrSQL & "        ELSE CAST(Price * Qty * ISNULL(TaxCategoryPercent,0) / 100.0 AS DECIMAL(38,15)) "
    StrSQL = StrSQL & "    END AS VatPercent , "
    StrSQL = StrSQL & "    CASE "
    StrSQL = StrSQL & "        WHEN ISNULL(VatValue,0) = 0 THEN 0 "
    StrSQL = StrSQL & "        ELSE ISNULL(TaxCategoryPercent,0) "
    StrSQL = StrSQL & "    END AS  Vatyo "
    StrSQL = StrSQL & "FROM tblEInvoice2 "
    StrSQL = StrSQL & "WHERE InvoiceID = '" & order_no & "'"


ElseIf docType = 5 Then
    StrSQL = " SELECT   top 1  IsNull(Remarks,'Note')   ItemName, Remarks as ItemNamee,ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, TblHandWages.TotalBvat as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.TblHandWages.TotalBvat  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.TblHandWages"
      StrSQL = StrSQL & " Where id = " & Transaction_ID & ""
ElseIf docType = 6 Then
    StrSQL = " SELECT   top 1  IsNull(Remarks,'Note')   ItemName, Remarks as ItemNamee,ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, TblHandWages.TotalBvat as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.TblHandWages.TotalBvat  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.TblHandWages"
      
      
    StrSQL = " SELECT  top 1 "
    StrSQL = StrSQL & "tblContractInsAllocationsDetails.ID"
    StrSQL = StrSQL & ",TblAqar.aqarname + ' ' + TblAqarDetai.unitno  as ItemName ,ItemWithOutVAT = 0"
    StrSQL = StrSQL & ",1 as ShowQty ,0 as PerforValue,tblContractInsAllocationsDetails.installValue  as showPrice"
    StrSQL = StrSQL & ", 0 as discountvalue,tblContractInsAllocationsDetails.ContNo"
    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.VATValue vatPercent,TblContract.FATYou as Vatyo "
'    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.HNoteSerial1"
'    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.installValue"
    StrSQL = StrSQL & "From TblAqarDetai "
    StrSQL = StrSQL & "INNER JOIN TblAqar "
    StrSQL = StrSQL & "    ON TblAqarDetai.Aqarid = TblAqar.Aqarid "
    StrSQL = StrSQL & "RIGHT OUTER JOIN TblContractInstallments "
    StrSQL = StrSQL & "INNER JOIN TblContract "
    StrSQL = StrSQL & "    ON TblContractInstallments.ContNo = TblContract.ContNo "
    StrSQL = StrSQL & "INNER JOIN tblContractInsAllocationsDetails "
    StrSQL = StrSQL & "    ON TblContractInstallments.InstallNo = tblContractInsAllocationsDetails.InstallNo "
    StrSQL = StrSQL & "        AND TblContract.ContNo = tblContractInsAllocationsDetails.ContNo "
    StrSQL = StrSQL & "    ON TblAqar.Aqarid = TblContract.Iqar "
    StrSQL = StrSQL & "        AND TblAqarDetai.Id = TblContract.UnitNo "


    StrSQL = StrSQL & " Where tblContractInsAllocationsDetails.id = " & Transaction_ID & ""
    
    
    If SystemOptions.CanUploadZakatOpt Then
'                StrSQL = " SELECT TOP 1 "
        StrSQL = ""
        
        ' œ«·… ’€Ì—… ÃÊ¯« SQL · ÕÊÌ· 0.15 ? 15 œÌ‰«„ÌﬂÌ«
        ' (Â  ﬂ » inline ﬂ‹ CASE)
        ' VAT% ··»‰œ: ·Ê ÷—«∆» «·»‰œ > 0 ‰«Œœ ‰”»… «·⁄ﬁœ° Ê≈·« 0
        
        ' ========= ≈ÌÃ«— =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ≈ÌÃ«—' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ≈ÌÃ«—' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchRentValueTaxes,0)=0 AND ISNULL(B.BatchRentValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchRentValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchRentValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchRentValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchRentValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= ﬂÂ—»«¡ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ﬂÂ—»«¡' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ﬂÂ—»«¡' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchElectricityValueTaxes,0)=0 AND ISNULL(B.BatchElectricityValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchElectricityValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchElectricityValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchElectricityValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchElectricityValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= „Ì«Â =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - „Ì«Â' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - „Ì«Â' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchWaterValueTaxes,0)=0 AND ISNULL(B.BatchWaterValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchWaterValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchWaterValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchWaterValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchWaterValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= Œœ„«  =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - Œœ„« ' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - Œœ„« ' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchServicesValueTaxes,0)=0 AND ISNULL(B.BatchServicesValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchServicesValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchServicesValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchServicesValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchServicesValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= «·”¯Û⁄Ì («·⁄„Ê·…) =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - «·”⁄Ì' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - «·”⁄Ì' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchCommissionValueTaxes,0)=0 AND ISNULL(B.BatchCommissionValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchCommissionValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchCommissionValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchCommissionValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchCommissionValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= €«“ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - €«“' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - €«“' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchGasValueTaxes,0)=0 AND ISNULL(B.BatchGasValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchGasValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchGasValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchGasValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchGasValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' =========  √„Ì‰ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' -  √„Ì‰' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' -  √„Ì‰' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchInsuranceValueTaxes,0)=0 AND ISNULL(B.BatchInsuranceValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchInsuranceValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchInsuranceValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchInsuranceValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchInsuranceValue,0) <> 0 "

    
    
    End If
    
     
End If
Dim mValueWithOutVat As Double
Dim ISVAT As Boolean
Dim classifiedTaxCategory2 As Double

    Set rs = New ADODB.Recordset
    If SystemOptions.CanUploadZakatOpt Then
       ' ConectionFirst
     '   Set cn2 = POSConnection
        rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Set cn2 = Cn
        rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
    End If
    PayableAmount = 0
    Totallinewithvat = 0
    Totallinewithoutvat = 0
    allowancechargeAmount = val(rs!discountvalue & "")
    If docType = 0 Then
     invline.Price.EncludingVat = SystemOptions.PriceWithVAT
   Else
        invline.Price.EncludingVat = False
    End If
   ISVAT = False
    
    allowancechargeAmount = 0
    Totallinewithvat = 0
    Totallinewithoutvat = 0
    TotalVat = 0

  ' =========================================
' BEGIN LOOP (REPLACEMENT)
' „›÷¯·  ’›Ì— «·„Ã«„Ì⁄ ﬁ»· «··Ê» (·Ê „‘ „ ’›¯—… »«·›⁄·):
' allowancechargeAmount = 0
' Totallinewithvat      = 0
' Totallinewithoutvat   = 0
' TotalVat              = 0

Dim lineQty As Double, linePrice As Double, lineDisc As Double
Dim lineVatAmt As Double, lineVatRate As Double, ratePct As Double
Dim lineTotal As Double, isItemExempt As Boolean
Dim useHiPrec As Boolean

useHiPrec = (docType = 10)

Do Until rs.EOF
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")

    ' 1) ﬁ—«¡… ¬„‰… ··ﬁÌ„ „‰ «·—ÌﬂÊ—œ” 
    lineQty = val("" & IIf(IsNull(rs("ShowQty").value), 0, rs("ShowQty").value))
    If useHiPrec Then
        linePrice = val("" & IIf(IsNull(rs("showPrice").value), 0, rs("showPrice").value))
    Else
        linePrice = Round(val("" & IIf(IsNull(rs("showPrice").value), 0, rs("showPrice").value)), 2)
    End If
    If docType = 1 Then
        lineDisc = 0
    Else
        lineDisc = Round(val("" & IIf(IsNull(rs("discountvalue").value), 0, rs("discountvalue").value)), 2)
    End If
    lineVatAmt = Round(val("" & IIf(IsNull(rs("vatPercent").value), 0, rs("vatPercent").value)), 2)      ' ﬁÌ„… «·÷—Ì»… („»·€)
    lineVatRate = Round(val("" & IIf(IsNull(rs("Vatyo").value), 0, rs("Vatyo").value)), 2)               ' ‰”»… «·÷—Ì»… %
    isItemExempt = (val("" & IIf(IsNull(rs("ItemWithOutVAT").value), 0, rs("ItemWithOutVAT").value)) <> 0)

    '  ÿ»Ì⁄ «·‰”»…: ·Ê 0.15 ‰Œ·ÌÂ« 15
    ratePct = IIf(lineVatRate > 1, lineVatRate, lineVatRate * 100)

    ' 2)  ⁄»∆… »Ì«‰«  «·”ÿ—
    invline.Item.Name = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
    invline.Price.PriceAmount = linePrice
    invline.InvoiceQuantity = lineQty

    ' 3)  ’‰Ì› ÷—Ì»Ì ·”ÿ—ﬂ
    If ratePct > 0 Then
        ' Œ«÷⁄
        ISVAT = True
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = ratePct
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = ratePct
    Else
        ' €Ì— Œ«÷⁄/„⁄›Ï/’›—Ì/Œ«—Ã «·‰ÿ«ﬁ
        If docType = 6 And isItemExempt Then
            ' ≈⁄›«¡ ⁄ﬁ«—Ì (”ﬂ‰Ì „À·«)
            invline.Item.classifiedTaxCategory.ID = "E"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-30"
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Real estate residential lease/sale"
        ElseIf InvoiceTypeCodename = "0100100" Then
            '  ’œÌ— ”·⁄ ’›—Ì…
            invline.Item.classifiedTaxCategory.ID = "Z"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-32"
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Export of goods"
        ElseIf InvoiceTypeCodename = "0100000" And docType2 <> 6 Then
            ' Œœ„«  „«·Ì… „⁄›«…
            invline.Item.classifiedTaxCategory.ID = "E"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
        ElseIf InvoiceTypeCodename = "0100000" And docType2 = 6 Then
            ' Œ«—Ã «·‰ÿ«ﬁ
            invline.Item.classifiedTaxCategory.ID = "O"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
               ' <-- «·≈÷«›… «·„ÿ·Ê»… · Ã«Ê“ BR-KSA-24 -->
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
        ElseIf InvoiceTypeCodename = "0200000" And ratePct = 0 Then
            ' „⁄›Ï (”Ì‰«—ÌÊ ›« Ê—… „»”ÿ… €Ì— Œ«÷⁄…)
            invline.Item.classifiedTaxCategory.ID = "E"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
        Else
            ' «› —«÷Ì: „⁄›Ï »œÊ‰ ”»» Œ«’
            invline.Item.classifiedTaxCategory.ID = "E"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
        End If
    End If
    
        ' ===== Prefer explicit category for WEB notes with no VAT =====
    If LCase$(mTableName) = "debitandcreditnotification" And ratePct = 0 Then
        ' ≈› —«÷Ì: Œ«—Ã ‰ÿ«ﬁ «·÷—Ì»… (OOS) „«·„ ⁄‰œﬂ ”Ì‰«—ÌÊ ¬Œ—
        invline.Item.classifiedTaxCategory.ID = "O"
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
    End If


    ' 4) ·Ê «·”⁄— = 0 «⁄ »—Â« Zero-rated (⁄Ì‰… „Ã«‰Ì…)
    If Round(linePrice, 2) = 0 Then
        invline.Item.classifiedTaxCategory.ID = "Z"
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-00"
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Free sample"
    End If

    ' 5)  ‰”Ìﬁ«  Õ”» docType
    If useHiPrec Then
        invline.InvoiceQuantity = Format(invline.InvoiceQuantity, ".##########")
        invline.Price.PriceAmount = Format(invline.Price.PriceAmount, ".##########")
        lineTotal = (val(invline.InvoiceQuantity) * val(invline.Price.PriceAmount)) - lineDisc
        lineTotal = val(Format(lineTotal, ".##########"))
    Else
        invline.InvoiceQuantity = Format(invline.InvoiceQuantity, ".##")
        invline.Price.PriceAmount = Format(invline.Price.PriceAmount, ".##")
        lineTotal = (val(invline.InvoiceQuantity) * val(invline.Price.PriceAmount)) - lineDisc
        lineTotal = val(Format(lineTotal, ".##"))
    End If

    ' 6)  Ã„Ì⁄ «·Œ’„° «·÷—Ì»…° «·≈Ã„«·Ì« 
    allowancechargeAmount = allowancechargeAmount + lineDisc          ' ·«  ’›¯—Â œ«Œ· «··Ê»
    If ratePct > 0 Then
        Totallinewithvat = Totallinewithvat + lineTotal               ' Ex-VAT lines subject to VAT
    Else
        Totallinewithoutvat = Totallinewithoutvat + lineTotal
    End If
    TotalVat = TotalVat + lineVatAmt                                   ' „Ã„Ê⁄ «·÷—Ì»… «·”ÿ—Ì…

    ' 7) ≈÷«›… «·”ÿ—
    inv.InvoiceLines.Add invline
    rs.MoveNext
Loop
' =========================================
' END LOOP (REPLACEMENT)



' ===== FINAL TOTALS (AFTER LOOP) =====
' „·«ÕŸ…: TotalVat  „  Ã„Ì⁄Â« ”ÿ—Ì« œ«Œ· «··Ê»° ·–·ﬂ ·« ‰⁄Ìœ Õ”«»Â« ﬂ‰”»… „‰ «·≈Ã„«·Ì.

'  ÊÕÌœ «·œﬁ… Õ”» ‰Ê⁄ «·„” ‰œ
If docType = 10 Then
    TotalVat = val(Format(TotalVat, ".##########"))
    Totallinewithvat = val(Format(Totallinewithvat, ".##########"))
    Totallinewithoutvat = val(Format(Totallinewithoutvat, ".##########"))
Else
    TotalVat = val(Format(TotalVat, ".##"))
    Totallinewithvat = val(Format(Totallinewithvat, ".##"))
    Totallinewithoutvat = val(Format(Totallinewithoutvat, ".##"))
End If

' ’«›Ì ﬁ»· «·÷—Ì»… + ÷—Ì»… + ’«›Ì €Ì— Œ«÷⁄ - Œ’Ê„«  ⁄·Ï „” ÊÏ «·›« Ê—… (≈‰ ÊıÃœ )
PayableAmount = Totallinewithvat + Totallinewithoutvat + TotalVat - mDiscountProject

' œ›⁄ ‰Â«∆Ì + ﬁÌ„… «·Œ’„ «·„Ã„¯⁄ (allowancecharge) »‰›” œﬁ… «·„” ‰œ
If docType = 10 Then
    inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmount, ".##########")
    allowancecharge.Amount = Format(allowancechargeAmount, "0.##########")
Else
    inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmount, ".##")
    allowancecharge.Amount = Format(allowancechargeAmount, "0.##")
End If

'    If dev.value = True Then
'        SendingMode = 0
'    ElseIf sim.value = True Then
'        SendingMode = 1
'    ElseIf Live.value = True Then
'        SendingMode = 2
'    End If
If mTableName = "tblEInvoice" Then
    mFieldIDName = "ID"
End If

SendingMode = SystemOptions.SendingMode

'here you can pass csid data
'„„ﬂ‰ ‰—”· »Ì«‰«  «· ceetificate „‰ «·œ« «»Ì“
If SendingMode <> 0 Then
 inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem  ' "" 'from file cert.pem
 inv.CSIDInfo.Privatekey = SystemOptions.Privatekey 'from file key.pem
'Dim SecretKey As String
SecretKey = SystemOptions.SecretKey  'from secret.txt file
End If


'WaelZakah
Set res = ubl.GenerateInvoiceXML(inv, App.path)




If mId(InvoiceTypeCodename, 1, 2) = "02" Then 'simple invoice

If res.IsValid = True Then
' Â–… «·‰ «∆Ã Ì—ÃÏ Õ›ŸÂ« ›Ï ﬁ«⁄œ… «·»Ì«‰«  «·Œ«’… »«·›« Ê—… ·ﬂÏ  ﬂÊ‰ „—Ã⁄ ·ﬂ„ ›Ì„« »⁄œ
InvoiceHash = res.InvoiceHash
SingedXML = res.SingedXML
EncodedInvoice = res.EncodedInvoice
 UUID = res.UUID
QRCode = res.QRCode
PIH = res.PIH
SingedXMLFileName = res.SingedXMLFileName

 

' after complete xml send it to zatca api
Dim apireqlogic As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
     Dim tokenresponse As Object
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
     Dim invrequestbody As Object
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
      Dim invoicereportingmodel As Object
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
     invrequestbody.invoice = res.EncodedInvoice
      invrequestbody.InvoiceHash = res.InvoiceHash
      invrequestbody.UUID = res.UUID
      
    
     
     If SendingMode = 0 Then 'dev
     apireqlogic.SetApiRequestLogic 0
         Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
         
    If tokenresponse.errorMessage = "" Then
        Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
 
                                    If invoicereportingmodel.IsSuccess = True Then
                      If invoicereportingmodel.StatusCode = 202 Then
                    '   MsgBox invoicereportingmodel.WarningMessage
                        
                                     WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                      '     str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                                      '       Cn2.Execute str
                       
                       
                      End If
                                 ' reporting status must be REPORTED
                                 ' MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                                  
                                 ErrorMessageS = invoicereportingmodel.ReportingStatus
                                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                                    zatcaStatus = 1
                                                    
                                                    str = " update " & mTableName & " set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                             cn2.Execute str
                                             updatedatatable val(docType), mTableName, mFieldIDName
                   
                                  Else
                                'MsgBox invoicereportingmodel.ErrorMessage
                                ErrorMessageS = invoicereportingmodel.errorMessage
                             savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                           '     str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                           '     Cn2.Execute str
                                End If
  Else
                    'MsgBox tokenresponse.ErrorMessage
                    ErrorMessageS = tokenresponse.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
 'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
 'Cn2.Execute str
 
 
 
                     
                        
                        
End If



     ElseIf SendingMode = 1 Then 'simu
      apireqlogic.SetApiRequestLogic 1
      Set invoicereportingmodel = apireqlogic.CallReportingAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                     If invoicereportingmodel.IsSuccess = True Then
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                             WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), mTableName, mFieldIDName
                        
                           '    str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           '                  Cn2.Execute str
                      End If
                     ' reporting status must be REPORTED
                  '    MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                      ErrorMessageS = invoicereportingmodel.ReportingStatus
                      savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                        zatcaStatus = 1
                              
                                        str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                  cn2.Execute str
 
                      updatedatatable val(docType), mTableName, mFieldIDName
                      Else
                  '  MsgBox invoicereportingmodel.ErrorMessage
                    ErrorMessageS = invoicereportingmodel.errorMessage
                      savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
             '       str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
             '        Cn2.Execute str
                     
                      
                    End If
        ElseIf SendingMode = 2 Then 'prood
        apireqlogic.SetApiRequestLogic 2
        Set invoicereportingmodel = apireqlogic.CallReportingAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                     If invoicereportingmodel.IsSuccess = True Then
                     ' reporting status must be REPORTED
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                                              WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                           '    str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           '                  Cn2.Execute str
                      End If
                  '    MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                      ErrorMessageS = invoicereportingmodel.ReportingStatus
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                         
                                                zatcaStatus = 1
                                           
                                        str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID


                                  cn2.Execute str
 
 updatedatatable val(docType), mTableName, mFieldIDName
                      Else
                          ErrorMessageS = errorMessage.ReportingStatus
               '     MsgBox invoicereportingmodel.ErrorMessage
                           'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                              savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                     'Cn2.Execute str
                    End If
        End If

   ' End If
  
   
Else
'MsgBox res.ErrorMessage
ErrorMessageS = res.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
       'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
       '              Cn2.Execute str
End If

Else





If res.IsValid = True Then
InvoiceHash = res.InvoiceHash
SingedXML = res.SingedXML
EncodedInvoice = res.EncodedInvoice
 UUID = res.UUID
QRCode = res.QRCode
PIH = res.PIH
SingedXMLFileName = res.SingedXMLFileName

 

' after complete xml send it to zatca api
'Dim apireqlogic As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
     'Dim tokenresponse As Object
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
     'Dim invrequestbody As Object
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
      'Dim invoicereportingmodel As Object
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
     Dim invoiceclearingmodel As Object
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")
     
     invrequestbody.invoice = res.EncodedInvoice
      invrequestbody.InvoiceHash = res.InvoiceHash
      invrequestbody.UUID = res.UUID
      
    ' Dim SendingMode As Integer
     
     If SendingMode = 0 Then 'dev
         Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
    If tokenresponse.errorMessage = "" Then
   
         Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                    If invoicereportingmodel.IsSuccess = True Then
                     
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                                                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                        
                             '  str = " update Transactions  set  warrningmessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                             '                Cn2.Execute str
                      End If
                      
                    
                     ' reporting status must be CLEARED
                   '   MsgBox invoicereportingmodel.ClearanceStatus 'CLEARED
                      ErrorMessageS = invoicereportingmodel.ClearanceStatus
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                       zatcaStatus = 1
                   
                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                        

                   cn2.Execute str 'WAELSALIM
                   updatedatatable val(docType), mTableName, mFieldIDName
         
 
                      Else
                  '  MsgBox invoicereportingmodel.ErrorMessage
                    ErrorMessageS = invoicereportingmodel.errorMessage
                       savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                           'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                     'Cn2.Execute str
                     
                    End If
  Else
'MsgBox tokenresponse.ErrorMessage
ErrorMessageS = tokenresponse.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
        'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
        '             Cn2.Execute str
End If
     ElseIf SendingMode = 1 Then 'simu
      
      apireqlogic.SetApiRequestLogic 1
     Call apireqlogic.SetToSaveClearedXMLFile(App.path, True)
      
      Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                If invoiceclearingmodel.IsSuccess = True Then
                     
                      If invoiceclearingmodel.StatusCode = 202 Then
                '       MsgBox invoiceclearingmodel.WarningMessage
                      '  str = " update Transactions  set  warrningmessage='" & invoiceclearingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           'MsgBox invoiceclearingmodel.WarningMessage
                            warrningmessage = invoiceclearingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(warrningmessage), , mTableName, mFieldIDName
                                         '    Cn2.Execute str
                      End If
                      
              '  MsgBox invoiceclearingmodel.ClearanceStatus 'Cleared
             ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, CStr(WarningMessage), CStr(ErrorMessageS), mTableName, mFieldIDName
                   'MsgBox invoiceclearingmodel.QRCode
                   QRCode = invoiceclearingmodel.QRCode
                   'MsgBox invoiceclearingmodel.ClearedInvoice
                                   zatcaStatus = 1
                   
                  
                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                 

               cn2.Execute str
 updatedatatable val(docType), mTableName, mFieldIDName
                  Else
                'MsgBox invoiceclearingmodel.ErrorMessage
                ErrorMessageS = invoiceclearingmodel.errorMessage
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                       'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                     'Cn2.Execute str
                End If
        ElseIf SendingMode = 2 Then 'prood
     apireqlogic.SetToSaveClearedXMLFile App.path, True
       apireqlogic.SetApiRequestLogic 2
        Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
     If invoiceclearingmodel.IsSuccess = True Then
                     
                      If invoiceclearingmodel.StatusCode = 202 Then
                    '   MsgBox invoiceclearingmodel.WarningMessage
                    WarningMessage = invoiceclearingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), mTableName, mFieldIDName
                        
                          '  str = " update Transactions  set  warrningmessage='" & invoiceclearingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                          '                   Cn2.Execute str
                      End If
                      
'MsgBox invoiceclearingmodel.ClearanceStatus 'Cleared
ErrorMessageS = invoiceclearingmodel.ClearanceStatus
savewarrninganderror Transaction_ID, , CStr(ErrorMessageS)
'   MsgBox invoiceclearingmodel.QRCode
   QRCode = invoiceclearingmodel.QRCode
   
  ' MsgBox invoiceclearingmodel.ClearedInvoice
                                 zatcaStatus = 1
                   'str = " update Transactions  set zatcaStatus  =1    where Transaction_ID =" & Transaction_ID

                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                        

               cn2.Execute str
               updatedatatable val(docType), mTableName, mFieldIDName
  Else
'MsgBox invoiceclearingmodel.ErrorMessage
ErrorMessageS = invoiceclearingmodel.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
         '      str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                 '    Cn2.Execute str
End If
        End If

   ' End If
  
   
Else
'MsgBox res.ErrorMessage
ErrorMessageS = res.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
     '          str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
     '                Cn2.Execute str
End If



End If




'
'Set res = ubl.GenerateInvoiceXML(inv, App.path)
'If res.IsValid = True Then
''MsgBox res.InvoiceHash
''MsgBox res.SingedXML
''MsgBox res.EncodedInvoice
''MsgBox res.UUID
''MsgBox res.QRCode
''MsgBox res.PIH
''MsgBox res.SingedXMLFileName
'
'InvoiceHash = res.InvoiceHash
'SingedXML = res.SingedXML
'EncodedInvoice = res.EncodedInvoice
'UUID = res.UUID
'QRCode = res.QRCode
'PIH = res.PIH
'SingedXMLFileName = res.SingedXMLFileName
'
'updatedatatable
'' after complete xml send it to zatca api
'Dim apireqlogic As ZatcaUBLXMLGenerator.ApiRequestLogic
'    Set apireqlogic = New ZatcaUBLXMLGenerator.ApiRequestLogic
'   Dim tokenresponse As ZatcaUBLXMLGenerator.ComplianceCsrResponse
'    Set tokenresponse = New ZatcaUBLXMLGenerator.ComplianceCsrResponse
'
'     Dim invrequestbody As ZatcaUBLXMLGenerator.InvoiceReportingRequest
'    Set invrequestbody = New ZatcaUBLXMLGenerator.InvoiceReportingRequest
'    Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
'    If tokenresponse.ErrorMessage = "" Then
'
'      invrequestbody.invoice = res.EncodedInvoice
'      invrequestbody.InvoiceHash = res.InvoiceHash
'      invrequestbody.UUID = res.UUID
'         Dim invoicereportingmodel As ZatcaUBLXMLGenerator.InvoiceReportingResponse
'         Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
' If invoicereportingmodel.ErrorMessage = "" Then
'ErrorMessageS = invoicereportingmodel.ReportingStatus
' ' reporting status must be CLEARED
' ' MsgBox invoicereportingmodel.ClearanceStatus 'CLEARED
'                  If invoicereportingmodel.ClearanceStatus = "CLEARED" Or invoicereportingmodel.ReportingStatus = "REPORTED" Then
'
'                   zatcaStatus = 1
'                   str = " update Transactions  set zatcaStatus  =1    where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'
'                If invoicereportingmodel.ClearanceStatus = "CLEARED" Then
'                ErrorMessageS = "CLEARED"
'                    str = " update Transactions  ErrorMessageS=  'CLEARED'   where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'                End If
'
'                       If invoicereportingmodel.ReportingStatus = "REPORTED" Then
'                ErrorMessageS = "REPORTED"
'                    str = " update Transactions  ErrorMessageS=  'REPORTED'   where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'                End If
'                  End If
'
'  Else
'
'MsgBox invoicereportingmodel.ErrorMessage
'ErrorMessageS = invoicereportingmodel.ErrorMessage
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'
'End If
'  Else
''MsgBox tokenresponse.ErrorMessage
'ErrorMessageS = "ıError in send " & tokenresponse.ErrorMessage & CHR(13) & "INV#" & ID
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'End If
'   ' End If
'
'
'Else
''MsgBox res.ErrorMessage
'ErrorMessageS = "ıError Before send in  " & res.ErrorMessage & CHR(13) & "INV#" & ID
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'End If


End Function

Public Function GenerateInvoice_ZatcaV2( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0) As Boolean

    On Error Resume Next: Err.Clear

    ' =======================
    '  Õ÷Ì— ﬂ«∆‰«  ZATCA SDK
    ' =======================
    Dim ubl As Object, inv As Object, res As Object
    Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
    Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
    Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

    'Dim cn2 As ADODB.Connection
    If SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
        Set cn2 = POSConnection
    Else
        Set cn2 = Cn
    End If

    If cschemeID = "" Then cschemeID = "CRN"
    If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
        fillmycompanydata ActivityTypeIdInv, branch_idInv
    Else
        fillmycompanydata
    End If

    ' =======================
    ' —√” «·›« Ê—…
    ' =======================
    inv.ID = ID
    If IssueDate = "" Then IssueDate = Date
    inv.IssueDate = Format$(IssueDate, "yyyy-mm-dd")

    If IssueTim = "" Then IssueTim = Now
    ' IMPORTANT: «·œﬁ«∆ﬁ nn („‘ MM)
    inv.IssueTime = Format$(IssueTim, "HH:nn:ss")

    inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)        ' 388 »Ì⁄ ñ 381/383 ≈‘⁄«—« 
    inv.InvoiceTypeCode.Name = InvoiceTypeCodename         ' „À· 0100000 √Ê 0200000

    If DocumentCurrencyCode = "" Or DocumentCurrencyCode = "SR" Then DocumentCurrencyCode = "SAR"
    If TaxCurrencyCode = "" Or TaxCurrencyCode = "SR" Then TaxCurrencyCode = "SAR"
    inv.DocumentCurrencyCode = DocumentCurrencyCode
    inv.TaxCurrencyCode = TaxCurrencyCode

    ' PIH «·√ŒÌ—
    inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
    ' ICV
    inv.AdditionalDocumentReferenceICV.UUID = uuidCounter

    '  Ê’Ì·/ ”·Ì„ («Œ Ì«—Ì ··›Ê« Ì— «·ﬁÌ«”Ì…)
    Dim ActualDeliveryDate As Date, LatestDeliveryDate As Date
    If ActualDeliveryDate = 0 Then ActualDeliveryDate = IssueDate
    If LatestDeliveryDate = 0 Then LatestDeliveryDate = IssueDate
    inv.Delivery.ActualDeliveryDate = Format$(ActualDeliveryDate, "yyyy-mm-dd")
    If val(InvoiceTypeCodeID) = 388 Then
        inv.Delivery.LatestDeliveryDate = Format$(LatestDeliveryDate, "yyyy-mm-dd")
    End If

    ' Ê”Ì·… «·œ›⁄
    Dim paymentmean As Object
    Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
    paymentmean.PaymentMeansCode = PaymentMeansCode
    paymentmean.PayeeFinancialAccount.ID = PayeeFinancialAccount
    paymentmean.PayeeFinancialAccount.paymentnote = paymentnote
    inv.PaymentMeans.Add paymentmean

    ' »Ì«‰«  «·„Ê—¯œ
    inv.SupplierParty.PartyIdentification.ID = cIdentificationID
    inv.SupplierParty.PartyIdentification.schemeID = cschemeID
    inv.SupplierParty.PostalAddress.StreetName = cStreetName
    inv.SupplierParty.PostalAddress.AdditionalStreetName = cAdditionalStreetName
    inv.SupplierParty.PostalAddress.BuildingNumber = cBuildingNumber
    inv.SupplierParty.PostalAddress.PlotIdentification = cPlotIdentification
    inv.SupplierParty.PostalAddress.CityName = cCityName
    inv.SupplierParty.PostalAddress.PostalZone = cPostalZone
    inv.SupplierParty.PostalAddress.CountrySubentity = cCountrySubentity
    inv.SupplierParty.PostalAddress.CitySubdivisionName = cCitySubdivisionName
    inv.SupplierParty.PostalAddress.Country.IdentificationCode = "SA"
    inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat)
    inv.SupplierParty.PartyTaxScheme.CompanyID = cCompanyID

    ' »Ì«‰«  «·⁄„Ì· (»»”«ÿ… ﬂ„« ›Ì ﬂÊœﬂ° „⁄ Defaults ·ÿÌ›…)
    Dim schemeID2 As String: schemeID2 = schemeID
    If mId(InvoiceTypeCodename, 1, 2) = "01" Then    ' ﬁÌ«”Ì…
        If Id700 <> "" Then
            inv.CustomerParty.PartyIdentification.ID = Id700
        Else
            inv.CustomerParty.PartyIdentification.ID = Identificationid
        End If
        If CompanyID = "" Then
            CompanyID = "300000000000003"
            schemeID2 = "OTH"
        End If
        inv.CustomerParty.PartyIdentification.schemeID = schemeID2

        If Identificationid = "N/A" Or Identificationid = "" Then
            If StreetName = "" Then StreetName = "N/A"
            If AdditionalStreetName = "" Then AdditionalStreetName = "N/A"
            If BuildingNumber = "" Then BuildingNumber = "N/A"
            If CityName = "" Then CityName = "N/A"
            If PostalZone = "" Then PostalZone = "00000"
            If CountrySubentity = "" Then CountrySubentity = "N/A"
            If CitySubdivisionName = "" Then CitySubdivisionName = "N/A"
        End If
        If IdentificationCode = "" Then IdentificationCode = "SA"
        inv.CustomerParty.PostalAddress.StreetName = StreetName
        inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName
        inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
        inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification
        inv.CustomerParty.PostalAddress.CityName = CityName
        inv.CustomerParty.PostalAddress.PostalZone = PostalZone
        inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity
        inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
        inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
        inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName)
        If Id700 <> "" Then
            inv.CustomerParty.PartyTaxScheme.CompanyID = ""
        Else
            inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID
        End If
    ElseIf mId(InvoiceTypeCodename, 1, 2) = "02" Then ' „»”ÿ…
        If CompanyID = "" Then
            CompanyID = "300000000000003"
            schemeID2 = "OTH"
        End If
        If Identificationid = "0" Then Identificationid = ""
        If Id700 <> "" Then
            inv.CustomerParty.PartyIdentification.ID = Id700
        Else
            inv.CustomerParty.PartyIdentification.ID = Identificationid
        End If
        inv.CustomerParty.PartyIdentification.schemeID = schemeID2
        If IdentificationCode = "" Then IdentificationCode = "SA"
        inv.CustomerParty.PostalAddress.StreetName = StreetName
        inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName
        inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
        inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification
        inv.CustomerParty.PostalAddress.CityName = CityName
        inv.CustomerParty.PostalAddress.PostalZone = PostalZone
        inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity
        inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
        inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
        inv.CustomerParty.PartyLegalEntity.RegistrationName = RegistrationName
        If Id700 <> "" Then
            inv.CustomerParty.PartyTaxScheme.CompanyID = ""
        Else
            inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID
        End If
    End If

    ' =======================
    ' AllowanceCharge ⁄·Ï „” ÊÏ «·›« Ê—…
    ' ‰⁄ „œ "Œ’„ »‰ÊœÌ ›ﬁÿ" -> —√” ›« Ê—… = ’›—
    ' =======================
    Dim allowancecharge As Object
    Set allowancecharge = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
    allowancecharge.Amount = "0.00"
    allowancecharge.AllowanceChargeReason = "Discount"
    If chkTaxExempt = 1 Or TaxCategoryPercent = 0 Then
        allowancecharge.taxCategory.ID = "E"
        allowancecharge.taxCategory.Percent = 0#
    Else
        allowancecharge.taxCategory.ID = "S"
        allowancecharge.taxCategory.Percent = 15#
    End If
    inv.allowanceCharges.Add allowancecharge

    ' =======================
    ' BillingReference: ›ﬁÿ ··≈‘⁄«—« 
    ' =======================
    If val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383 Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = val(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If

    ' =======================
    ' ﬁ—«¡… »‰Êœ «·›« Ê—… (Transactions) docType=0
    ' =======================
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim Percetage As Double

    ' Â‰” Œœ„ «·«” ⁄·«„ «·„‰÷»ÿ «··Ì ÌÕ”»:
    ' totalPrice, totalPriceAfterDiscount, showPrice (√”«” »œÊ‰ «·÷—Ì»… »⁄œ «·Œ’„), vatPercent („»·€ «·÷—Ì»… »⁄œ «·Œ’„), Vatyo («·‰”»…)
    If docType = 0 Then
      StrSQL = ""
StrSQL = StrSQL & "SELECT"
StrSQL = StrSQL & "  I.ItemName,"
StrSQL = StrSQL & "  I.ItemNamee,"
StrSQL = StrSQL & "  ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT,"
StrSQL = StrSQL & "  TD.ShowQty,"
StrSQL = StrSQL & "  PerforValue = 0,"

' Œ’„ «·”ÿ— «·ÕﬁÌﬁÌ (»‰ÊœÌ): ‰Ã„⁄ ItemDiscount + TotalDiscountPerLine
StrSQL = StrSQL & "  (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)) AS discountvalue,"

' Œ’„ —√” «·›« Ê—… (·Ê ⁄‰œﬂ) ﬂ„« ÂÊ ··⁄—÷
StrSQL = StrSQL & "  T.Trans_Discount AS discountvalue2,"
StrSQL = StrSQL & "  T.Trans_DiscountType,"

If SystemOptions.PriceWithVAT Then
    ' Â‰” Œœ„ ‰”»… «·÷—Ì»… ·ﬂ· ”ÿ— ·Ê „ «Õ… TD.Vatyo Ê≈·« Percetage «·⁄«„…
    PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage

    ' „⁄œ· «·÷—Ì»… ··”ÿ— (?)
    StrSQL = StrSQL & "  RatePct = CASE"
    StrSQL = StrSQL & "              WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "                   THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "                   ELSE 0"
    StrSQL = StrSQL & "            END,"

    ' ≈Ã„«·Ì «·”ÿ— (”⁄— „Œ“‰ ‘«„· ÷—Ì»… * ﬂ„Ì…) ﬁ»· «·Œ’„
    StrSQL = StrSQL & "  ROUND(TD.showPrice * TD.ShowQty, 4) AS totalPrice,"

    ' ≈Ã„«·Ì «·”ÿ— »⁄œ «·Œ’„ (·”Â ‘«„· ÷—Ì»…)
    StrSQL = StrSQL & "  GrossAfterDisc = ROUND((TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "                       - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)), 4),"

    ' «·√”«” »œÊ‰ ÷—Ì»… »⁄œ «·Œ’„ (··Õ”«»« )
    StrSQL = StrSQL & "  NetAfterDisc = ROUND("
    StrSQL = StrSQL & "     ((TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "        - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)))"
    StrSQL = StrSQL & "     / (1 + ((CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "                   THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "                   ELSE 0 END) / 100.0)), 4),"

    ' ”⁄— «·ÊÕœ… «·’ÕÌÕ ·≈—”«· UBL: √”«” ﬁ»· «·Œ’„ Ê»œÊ‰ ÷—Ì»…
    StrSQL = StrSQL & "  showPrice = ROUND("
    StrSQL = StrSQL & "     TD.showPrice / (1 + ((CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "                                  THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "                                  ELSE 0 END) / 100.0))"
    StrSQL = StrSQL & "  , 4),"

    ' ≈Ã„«·Ì «·”ÿ— ì»⁄œ «·Œ’„î (·· Ê«›ﬁ „⁄ «”„ﬂ «·ﬁœÌ„) = NetAfterDisc
    StrSQL = StrSQL & "  totalPriceAfterDiscount = "
    StrSQL = StrSQL & "     ROUND(((TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "            - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)))"
    StrSQL = StrSQL & "           / (1 + ((CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "                          THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "                          ELSE 0 END) / 100.0))"
    StrSQL = StrSQL & "     , 4),"

    ' „»·€ «·÷—Ì»… »⁄œ «·Œ’„ (‘«„· - €Ì— ‘«„·)
    StrSQL = StrSQL & "  vatPercent = ROUND("
    StrSQL = StrSQL & "     ((TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "       - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)))"
    StrSQL = StrSQL & "     - ("
    StrSQL = StrSQL & "        ((TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "           - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)))"
    StrSQL = StrSQL & "        / (1 + ((CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "                       THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "                       ELSE 0 END) / 100.0))"
    StrSQL = StrSQL & "       )"
    StrSQL = StrSQL & "  , 4),"

    ' ‰”»… «·÷—Ì»… ··⁄—÷/«·≈—”«·
    StrSQL = StrSQL & "  Vatyo = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0"
    StrSQL = StrSQL & "               THEN CAST(NULLIF(ISNULL(TD.Vatyo,0),0) AS DECIMAL(18,2))"
    StrSQL = StrSQL & "               ELSE 0 END"
Else
    ' «·”⁄— „Œ“¯‰ €Ì— ‘«„· «·÷—Ì»…
    ' ≈Ã„«·Ì «·”ÿ— ﬁ»· «·Œ’„ (€Ì— ‘«„·)
    StrSQL = StrSQL & "  totalPrice = ROUND(TD.showPrice * TD.ShowQty, 4),"

    ' ≈Ã„«·Ì «·”ÿ— »⁄œ «·Œ’„ (€Ì— ‘«„·)
    StrSQL = StrSQL & "  totalPriceAfterDiscount = ROUND("
    StrSQL = StrSQL & "      (TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "      - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0)), 4),"

    ' ”⁄— «·ÊÕœ… ·≈—”«· UBL: ﬁ»· «·Œ’„ (€Ì— ‘«„·)
    StrSQL = StrSQL & "  showPrice = ROUND(TD.showPrice, 4),"

    ' „»·€ «·÷—Ì»… »⁄œ «·Œ’„
    StrSQL = StrSQL & "  vatPercent = ROUND("
    StrSQL = StrSQL & "     ("
    StrSQL = StrSQL & "       (TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "       - (ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0))"
    StrSQL = StrSQL & "     ) * (CAST(CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END AS DECIMAL(18,2)) / 100.0)"
    StrSQL = StrSQL & "  , 4),"

    ' ‰”»… «·÷—Ì»…
    StrSQL = StrSQL & "  Vatyo = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END"
End If

StrSQL = StrSQL & " FROM dbo.Transactions AS T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID"
StrSQL = StrSQL & " WHERE T.Transaction_ID = " & Transaction_ID

    Else
        ' TODO: œ⁄„ »«ﬁÌ «·√‰Ê«⁄ ·«Õﬁ«
        GenerateInvoice_ZatcaV2 = False
        Exit Function
    End If

    rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF And rs.BOF Then
        GenerateInvoice_ZatcaV2 = False
        Exit Function
    End If

    ' =======================
    '  Ã„Ì⁄ «·”ÿÊ— Ê«·≈Ã„«·Ì« 
    ' =======================
    Dim Totallinewithvat As Double, Totallinewithoutvat As Double, TotalVat As Double
    Totallinewithvat = 0: Totallinewithoutvat = 0: TotalVat = 0

    Dim invline As Object
    Dim lineQty As Double, linePrice As Double, lineDisc As Double
    Dim lineVatAmt As Double, lineVatRate As Double, ratePct As Double
    Dim lineTotal As Double, isItemExempt As Boolean

    Do Until rs.EOF
        Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")

        lineQty = val("" & Nz(rs("ShowQty").value, 0))
        linePrice = Round(val("" & Nz(rs("showPrice").value, 0)), 4) ' œﬁ… √⁄·Ï ‘ÊÌ…
        lineDisc = Round(val("" & Nz(rs("discountvalue").value, 0)), 4) ' Œ’„ »‰ÊœÌ ›ﬁÿ
        lineVatAmt = Round(val("" & Nz(rs("vatPercent").value, 0)), 4)
        lineVatRate = Round(val("" & Nz(rs("Vatyo").value, 0)), 2)
        isItemExempt = (val("" & Nz(rs("ItemWithOutVAT").value, 0)) <> 0)

        ratePct = IIf(lineVatRate > 1, lineVatRate, lineVatRate * 100)

        invline.Item.Name = NzText(rs("ItemName").value)
        invline.Price.PriceAmount = linePrice
        invline.Price.EncludingVat = SystemOptions.PriceWithVAT
        invline.InvoiceQuantity = Format$(lineQty, "0.##########")

        '  ’‰Ì› ÷—Ì»Ì
        If ratePct > 0 Then
            invline.Item.classifiedTaxCategory.ID = "S"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
            invline.Item.classifiedTaxCategory.Percent = ratePct
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = ratePct
        Else
            ' „⁄›Ï/’›—/Œ«—Ã «·‰ÿ«ﬁ Õ”» ”Ì‰«—ÌÊÂ« ﬂ
            invline.Item.classifiedTaxCategory.ID = "E"
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
            invline.Item.classifiedTaxCategory.Percent = 0
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
        End If

        ' ≈Ã„«·Ì «·”ÿ— ﬁ»· «·÷—Ì»… »⁄œ «·Œ’„ «·»‰ÊœÌ
        lineTotal = (lineQty * linePrice) - lineDisc
        lineTotal = val(Format$(lineTotal, "0.##########"))

        If ratePct > 0 Then
            Totallinewithvat = Totallinewithvat + lineTotal
        Else
            Totallinewithoutvat = Totallinewithoutvat + lineTotal
        End If
        TotalVat = TotalVat + lineVatAmt

        inv.InvoiceLines.Add invline
        rs.MoveNext
    Loop
    rs.Close

    ' =======================
    ' LegalMonetaryTotal („ ¯”ﬁ)
    ' =======================
    Dim sumLinesBeforeVAT As Double
    sumLinesBeforeVAT = Totallinewithvat + Totallinewithoutvat

    Dim allowance As Double: allowance = 0 ' Œ’„ —√” ›« Ê—… = ’›— (»‰ÊœÌ ›ﬁÿ)
    Dim taxExclusive As Double: taxExclusive = sumLinesBeforeVAT - allowance
    Dim TaxInclusive As Double: TaxInclusive = taxExclusive + TotalVat

    inv.LegalMonetaryTotal.LineExtensionAmount = Format$(sumLinesBeforeVAT, "0.00")
    inv.LegalMonetaryTotal.AllowanceTotalAmount = Format$(allowance, "0.00")
    inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
    inv.LegalMonetaryTotal.TaxExclusiveAmount = Format$(taxExclusive, "0.00")
    inv.LegalMonetaryTotal.TaxInclusiveAmount = Format$(TaxInclusive, "0.00")
    inv.LegalMonetaryTotal.PayableAmount = Format$(TaxInclusive, "0.00")
    'inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00" ' «Œ Ì«—Ì

    ' =======================
    ' «· ÊﬁÌ⁄ Ê«·≈—”«·
    ' =======================
    Dim SendingMode As Integer
    SendingMode = SystemOptions.SendingMode

    If SendingMode <> 0 Then
        inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
        inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
    End If
    Dim SecretKey As String: SecretKey = SystemOptions.SecretKey

    Set res = ubl.GenerateInvoiceXML(inv, App.path)

    If res.IsValid = True Then
        ' Œ“¯‰ ‰Ê« Ã «·≈‰‘«¡
        InvoiceHash = res.InvoiceHash
        SingedXML = res.SingedXML
        EncodedInvoice = res.EncodedInvoice
        UUID = res.UUID
        QRCode = res.QRCode
        PIH = res.PIH
        SingedXMLFileName = res.SingedXMLFileName

        ' ≈—”«· Õ”» «·„Êœ
        Dim apireqlogic As Object
        Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")

        Dim tokenresponse As Object
        Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")

        Dim invrequestbody As Object
        Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
        invrequestbody.invoice = res.EncodedInvoice
        invrequestbody.InvoiceHash = res.InvoiceHash
        invrequestbody.UUID = res.UUID

        Dim invoicereportingmodel As Object
        Dim invoiceclearingmodel As Object

        If mTableName = "tblEInvoice" Then mFieldIDName = "ID"

        If SendingMode = 0 Then
            ' DEV: Compliance
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus = 1 WHERE " & mFieldIDName & " = " & Transaction_ID
                    updatedatatable CInt(docType), mTableName, mFieldIDName
                    GenerateInvoice_ZatcaV2 = True
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    GenerateInvoice_ZatcaV2 = False
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                GenerateInvoice_ZatcaV2 = False
            End If

        ElseIf SendingMode = 1 Then
            ' SIMULATION: Clearance
            apireqlogic.SetApiRequestLogic 1
            Call apireqlogic.SetToSaveClearedXMLFile(App.path, True)
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus = 1 WHERE " & mFieldIDName & " = " & Transaction_ID
                updatedatatable CInt(docType), mTableName, mFieldIDName
                GenerateInvoice_ZatcaV2 = True
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                GenerateInvoice_ZatcaV2 = False
            End If

        ElseIf SendingMode = 2 Then
            ' PROD: Clearance
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            apireqlogic.SetApiRequestLogic 2
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus = 1 WHERE " & mFieldIDName & " = " & Transaction_ID
                updatedatatable CInt(docType), mTableName, mFieldIDName
                GenerateInvoice_ZatcaV2 = True
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                GenerateInvoice_ZatcaV2 = False
            End If
        End If

    Else
        ErrorMessageS = res.errorMessage
        savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
        GenerateInvoice_ZatcaV2 = False
    End If
End Function


Public Function GenerateInvoice_ZatcaV3Last( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "Transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0 _
) As Boolean
On Error Resume Next

' =====================  Õ÷Ì— «·‹ UBL/Invoice =====================
Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

' ---- « ’«· (ﬂ„« ÂÊ »œÊ‰  ⁄œÌ·) ----
If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If

' ---- »Ì«‰«  «·‘—ﬂ…/«·›—⁄ ----
If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

' ---------------------- BillingReference (CN/DN only) ----------------------
If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If
End If

' ---------------------- ‰Ê⁄ «·›« Ê—… ----------------------
inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
inv.InvoiceTypeCode.Name = InvoiceTypeCodename

' ---------------------- ICV / UUID / PIH ----------------------
inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)

Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
icvOnlyDigits = ""
icvStr = Trim$(uuidCounter & "")
If icvStr = "" Then icvStr = "1"
For i = 1 To Len(icvStr)
    Ch = mId$(icvStr, i, 1)
    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
Next i
If icvOnlyDigits = "" Then icvOnlyDigits = "1"
inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits

inv.ID = ID

' ---------------------- Issue Date & Time ----------------------
If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")

If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
' „Â„: mm = œﬁ«∆ﬁ / ss = ÀÊ«‰Ì
inv.IssueTime = Format(IssueTim, "HH:mm:ss")

' ---------------------- Delivery Dates ----------------------
If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
End If

' ---------------------- Payment Means ----------------------
Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
    paymentmean.InstructionNote = "Returned items"
End If
If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
End If
If Len(Trim$(paymentnote & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
End If
inv.PaymentMeans.Add paymentmean

' ---------------------- Supplier ----------------------
inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
inv.SupplierParty.PartyIdentification.schemeID = cschemeID

inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
Dim supplierCountry As String
supplierCountry = Trim$(cIdentificationCode & "")
If supplierCountry = "" Then supplierCountry = "SA"
inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry
inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")

' ---------------------- Customer (Standard/Simplified) ----------------------
Dim isStandard As Boolean
isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")

Dim isSimplified As Boolean
isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
               And (LCase$(mTableName) = "tblEinvoice" _
                 Or LCase$(mTableName) = "propertyduebatchdetail" _
                 Or LCase$(mTableName) = "debitandcreditnotification")

If isStandard Or isSimplified Then
    Dim CUSTID As String
    If Len(Trim$(Id700 & "")) > 0 Then
        CUSTID = Trim$(Id700 & "")
    Else
        CUSTID = Trim$(Identificationid & "")
    End If
    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
        CUSTID = "300000000000003": schemeID = "OTH"
    End If

    inv.CustomerParty.PartyIdentification.ID = CUSTID
    If Len(Trim$(schemeID & "")) = 0 Then schemeID = "OTH"
    inv.CustomerParty.PartyIdentification.schemeID = schemeID

    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
    cAddStreet = Trim$(AdditionalStreetName & "")
    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
    cSubent = Trim$(CountrySubentity & "")
    cSubCity = Trim$(CitySubdivisionName & "")
    cCountry = Trim$(IdentificationCode & "")
    If Len(cCountry) = 0 Then cCountry = "SA"

    inv.CustomerParty.PostalAddress.StreetName = cStreet
    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
    inv.CustomerParty.PostalAddress.CityName = cCity
    inv.CustomerParty.PostalAddress.PostalZone = cPost
    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry

    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
    If Len(Trim$(Id700 & "")) > 0 Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
    End If
End If

' =========================================================
' ============   «·»‰Êœ +  Ã„Ì⁄ «·Œ’Ê„«    ===============
' =========================================================
Dim StrSQL As String
StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & " t.Trans_Discount,t.Trans_DiscountType, I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  LineDiscRaw = ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0), "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID




StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  TD.ItemDiscountType, "
StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
Dim SumExNoDisc As Double                 ' ≈Ã„«·Ì «·√”«” ﬁ»· √Ì Œ’„ (Ex-VAT)
Dim SumTaxableExNoDisc As Double          ' ≈Ã„«·Ì «·√”«” «·Œ«÷⁄ ﬁ»· «·Œ’„
Dim SumExemptExNoDisc As Double           ' ≈Ã„«·Ì «·√”«” «·„⁄›Ï ﬁ»· «·Œ’„
Dim SumLineDiscEx As Double               ' „Ã„Ê⁄ Œ’„ «·”ÿÊ— »⁄œ  ÕÊÌ·Â Ex-VAT

Do Until rs.EOF
    Dim lineEx As Double
    Dim Rate As Double
    Dim Qty As Double
    Dim basePrice As Double

    Dim discType As Long
    Dim discPercent As Double
    Dim discValue As Double
    Dim discRaw As Double
    Dim lineDiscEx As Double
    Dim lineBaseForPercent As Double

    ' ﬁÌ„ «·”ÿ— «·√”«”Ì…
    lineEx = val(rs!LineEx_NoDisc & "")   ' ﬁÌ„… «·”ÿ— ﬁ»· «·Œ’„ Ex-VAT
    Rate = val(rs!vatRate & "")
    Qty = val(rs!ShowQty & "")
    basePrice = val(rs!BasePrice_Ex & "")

    ' »Ì«‰«  «·Œ’„
    discType = val(rs!ItemDiscountType & "")
    discPercent = val(rs!ItemDiscount & "")   '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ‰”»…
    discValue = val(rs!ItemDiscount & "")     '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ﬁÌ„…
    discRaw = 0

    ' ===== Õ”«» «·Œ’„ «·„Œ“‰ ⁄·Ï «·”ÿ— (»«·√”⁄«— «· Ì √œŒ·Â« «·„” Œœ„) =====
    Select Case discType
        Case 3      ' Œ’„ ‰”»… %
            ' ·Ê «·√”⁄«— »«·÷—Ì»… ? ‰”» Â „‰ ﬁÌ„… «·”ÿ— »«·÷—Ì»…
            ' ·Ê «·√”⁄«— »œÊ‰ ÷—Ì»… ? ‰”» Â „‰ Ex-VAT
            If SystemOptions.PriceWithVAT And Rate > 0 Then
                ' lineEx ÂÊ Ex-VAT° ‰ÕÊ·Â ≈·Ï ‘«„· ÷—Ì»…
                lineBaseForPercent = lineEx * (1 + Rate / 100#)
            Else
                lineBaseForPercent = lineEx
            End If

            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) _
                      + val(rs!TotalDiscountPerLine & "")

        Case 2      ' Œ’„ ﬁÌ„…
            discRaw = discValue + val(rs!TotalDiscountPerLine & "")

        Case Else   ' „›Ì‘ ‰Ê⁄ Œ’„ „Õœœ ? ‰ﬂ ›Ì »‹ TotalDiscountPerLine ·Ê „ÊÃÊœ
            discRaw = val(rs!TotalDiscountPerLine & "")
    End Select

    ' =====  ÕÊÌ· «·Œ’„ ≈·Ï Ex-VAT =====
    If SystemOptions.PriceWithVAT And Rate > 0 Then
        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
    Else
        lineDiscEx = Round(discRaw, 6)
    End If

    ' =====  Ã„Ì⁄ «·√”«”«  ﬁ»· «·Œ’„ =====
    SumExNoDisc = SumExNoDisc + lineEx

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
    Else
        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
    End If

    ' =====  Ã„Ì⁄ Œ’Ê„«  «·”ÿÊ— ﬂ‹ "Œ’„ ﬂ·Ì" ⁄·Ï «·›« Ê—… =====
    SumLineDiscEx = SumLineDiscEx + lineDiscEx

    ' ===== ≈‰‘«¡ ”ÿ— UBL »·« Œ’„ ”ÿ—Ì =====
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    invline.Price.EncludingVat = False
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = rs!ItemName & ""

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = Rate
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = Rate
    Else
        invline.Item.classifiedTaxCategory.ID = "E"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
    End If

    ' ·« ‰÷Ì› AllowanceCharge ⁄·Ï „” ÊÏ «·”ÿ— ≈ÿ·«ﬁ«
    inv.InvoiceLines.Add invline

    rs.MoveNext
Loop

rs.Close

' =========================================================
' ==============  Œ’„ „ÊÕ¯œ (”ÿÊ— + ÂÌœ—)  ================
' =========================================================
Dim HeaderDiscEx As Double

' SumLineDiscEx Ã«Ì „‰ «··Ê» : „Ã„Ê⁄ Œ’Ê„«  «·”ÿÊ— »⁄œ  ÕÊÌ·Â« Ex-VAT
HeaderDiscEx = Round(SumLineDiscEx, 6)

' ·Ê ⁄‰œﬂ Œ’„ —√” ›« Ê—… (Trans_Discount) Ê —Ìœ ÷„¯Â ﬂ„»·€ (Ex-VAT)
If val(Trans_Discount & "") <> 0 Then
    HeaderDiscEx = HeaderDiscEx + Round(val(Trans_Discount & ""), 6)
End If

If HeaderDiscEx < 0 Then HeaderDiscEx = 0

'  Ê“Ì⁄ œ«Œ·Ì (‰ˆ”»Ì) ·„⁄—›…  √ÀÌ— «·Œ’„ ⁄·Ï «·Œ«÷⁄/«·„⁄›Ï ·Õ”«» VAT
Dim BaseTotal As Double: BaseTotal = SumExNoDisc
Dim Allow_S As Double, Allow_E As Double

If BaseTotal > 0 Then
    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 6)
    Allow_E = Round(HeaderDiscEx - Allow_S, 6)
Else
    Allow_S = 0
    Allow_E = HeaderDiscEx
End If

' --------- √÷› AllowanceCharge ··ÂÌœ— (Ê«Õœ √Ê «À‰Ì‰) ---------
If HeaderDiscEx > 0# Then

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S
    If Allow_S > 0# Then
        Dim allowanceS As Object
        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceS.ChargeIndicator = False
        allowanceS.Amount = Format(Allow_S, "0.##")
        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceS.taxCategory.ID = "S"
        allowanceS.taxCategory.Percent = 15#
        inv.allowanceCharges.Add allowanceS
    End If

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·„⁄›«… E
    If Allow_E > 0# Then
        Dim allowanceE As Object
        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceE.ChargeIndicator = False
        allowanceE.Amount = Format(Allow_E, "0.##")
        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceE.taxCategory.ID = "E"
        allowanceE.taxCategory.Percent = 0#
        inv.allowanceCharges.Add allowanceE
    End If
End If

' =========================================================
' ====================   «·≈Ã„«·Ì«    =====================
' =========================================================

' LineExtensionAmount = √”«” «·»‰Êœ ﬁ»· √Ì Œ’„ Ê÷—Ì»… (Ex-VAT ﬁ»· «·Œ’„)
inv.LegalMonetaryTotal.LineExtensionAmount = Format(Round(SumExNoDisc, 2), "0.##")

' AllowanceTotalAmount = ≈Ã„«·Ì «·Œ’„ «·„ÊÕ¯œ (”ÿÊ— + ÂÌœ—) Ex-VAT
inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(Round(HeaderDiscEx, 2), "0.##")

' TaxExclusiveAmount = LineExtension - Allowance (Ex-VAT »⁄œ «·Œ’„)
Dim TaxExclusiveAfterDisc As Double
TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 6)
If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0
inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(Round(TaxExclusiveAfterDisc, 2), "0.##")

' VAT = ⁄·Ï «·Ã“¡ «·Œ«÷⁄ »⁄œ «·Œ’„
Dim TaxableAfterDisc As Double
TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 6)
If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0

Dim TotalVat As Double
TotalVat = Round(TaxableAfterDisc * 0.15, 2)
inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")

' TaxInclusive = TaxExclusive + VAT
Dim TaxInclusive As Double
TaxInclusive = Round(TaxExclusiveAfterDisc + TotalVat, 2)
inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(TaxInclusive, "0.##")

' «·ÕﬁÊ· «·√Œ—Ï
inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"

If val(InvoiceTypeCodeID) = 388 Then
    inv.LegalMonetaryTotal.PrepaidAmount = Format(val(PrepaidAmount & ""), "0.##")
Else
    inv.LegalMonetaryTotal.PrepaidAmount = "0.00"
End If

' Payable = TaxInclusive - Prepaid (·Ê ›Ì „œ›Ê⁄«  „”»ﬁ…)
Dim PayableAmountCalc As Double
PayableAmountCalc = TaxInclusive - val(inv.LegalMonetaryTotal.PrepaidAmount)
If PayableAmountCalc < 0 Then PayableAmountCalc = 0
inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmountCalc, "0.##")


' =========================================================
' ==================   «· ÊﬁÌ⁄ Ê«·≈—”«·   =================
' =========================================================
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey

If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)

If res.IsValid = True Then
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    Dim apireqlogic As Object, tokenresponse As Object
    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")

    invrequestbody.invoice = res.EncodedInvoice
    invrequestbody.InvoiceHash = res.InvoiceHash
    invrequestbody.UUID = res.UUID

    Select Case SendingMode
        Case 0 ' DEV (Compliance)
            apireqlogic.SetApiRequestLogic 0
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                    updatedatatable val(docType), mTableName, mFieldIDName
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 1 ' SIMU (Reporting/Clearance)
            apireqlogic.SetApiRequestLogic 1
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 2 ' PROD (Clearance)
            apireqlogic.SetApiRequestLogic 2
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If
    End Select

    GenerateInvoice_ZatcaV3Last = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV3Last = False
End If

End Function

'Public Function GenerateInvoice_ZatcaV3( _
'    Optional ByVal docType As Integer = 0, _
'    Optional ByVal mTableName As String = "Transactions", _
'    Optional ByVal mFieldIDName As String = "Transaction_ID", _
'    Optional ByVal ActivityTypeIdInv As Integer = 0, _
'    Optional ByVal branch_idInv As Integer = 0 _
') As Boolean
'On Error Resume Next
'
'    ' === ÿ•ÿ∞ÿß ÿßŸÑÿ£ÿ≥ÿπÿßÿ± ÿ¥ÿßŸÖŸÑÿ© ÿßŸÑÿ∂ÿ±Ÿäÿ®ÿ©ÿå ÿßÿ≥ÿ™ÿÆÿØŸÖ ÿßŸÑÿØÿßŸÑÿ© ÿßŸÑŸÇÿØŸäŸÖÿ© ===
'    If SystemOptions.PriceWithVAT Then
'        Call generateInvoice(docType, mTableName, mFieldIDName, ActivityTypeIdInv, branch_idInv)
'        GenerateInvoice_ZatcaV3 = True
'        Exit Function
'    End If
'
'' ===================== ÿ™ÿ≠ÿ∂Ÿäÿ± ÿßŸÑŸÄ UBL/Invoice =====================
'Dim ubl As Object, inv As Object, res As Object
'Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
'Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
'Set res = CreateObject("ZatcaUBLXMLGenerator.Result")
'
'' ---- ÿßÿ™ÿµÿßŸÑ ŸÇÿßÿπÿØÿ© ÿßŸÑÿ®ŸäÿßŸÜÿßÿ™ ----
'If SystemOptions.CanUploadZakatOpt Then
'    ConectionFirst
'    Set cn2 = POSConnection
'Else
'    Set cn2 = Cn
'End If
'
'' ---- ÿ®ŸäÿßŸÜÿßÿ™ ÿßŸÑÿ¥ÿ±ŸÉÿ©/ÿßŸÑŸÅÿ±ÿπ ----
'If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
'    fillmycompanydata ActivityTypeIdInv, branch_idInv
'Else
'    fillmycompanydata
'End If
'
'' ---------------------- BillingReference (CN/DN only) ----------------------
'If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
'    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
'        Dim InvoiceDocumentReference As Object
'        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
'        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
'        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
'    End If
'End If
'
'' ---------------------- ŸÜŸàÿπ ÿßŸÑŸÅÿßÿ™Ÿàÿ±ÿ© ----------------------
'inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
'inv.InvoiceTypeCode.Name = InvoiceTypeCodename
'
'' ---------------------- ÿßŸÑÿπŸÖŸÑÿ© (ŸÖŸáŸÖ - ŸÉÿßŸÜ ŸÜÿßŸÇÿµ!) ----------------------
'If DocumentCurrencyCode = "" Or DocumentCurrencyCode = "SR" Then DocumentCurrencyCode = "SAR"
'If TaxCurrencyCode = "" Or TaxCurrencyCode = "SR" Then TaxCurrencyCode = "SAR"
'inv.DocumentCurrencyCode = DocumentCurrencyCode
'inv.TaxCurrencyCode = TaxCurrencyCode
'
'' ---------------------- ICV / UUID / PIH ----------------------
'inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
'
'Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
'icvOnlyDigits = ""
'icvStr = Trim$(uuidCounter & "")
'If icvStr = "" Then icvStr = "1"
'For i = 1 To Len(icvStr)
'    Ch = mId$(icvStr, i, 1)
'    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
'Next i
'If icvOnlyDigits = "" Then icvOnlyDigits = "1"
'inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits
'
'inv.ID = ID
'
'' ---------------------- Issue Date & Time ----------------------
'If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
'inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")
'
'If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
'inv.IssueTime = Format(IssueTim, "HH:mm:ss")
'
'' ---------------------- Delivery Dates ----------------------
'If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
'If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
'inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
'If val(InvoiceTypeCodeID) = 388 Then
'    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
'End If
'
'' ---------------------- Payment Means ----------------------
'Dim paymentmean As Object
'Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
'paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
'If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
'    paymentmean.InstructionNote = "Returned items"
'End If
'If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
'    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
'End If
'If Len(Trim$(paymentnote & "")) > 0 Then
'    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
'End If
'inv.PaymentMeans.Add paymentmean
'
'' ---------------------- Supplier ----------------------
'inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
'If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
'inv.SupplierParty.PartyIdentification.schemeID = cschemeID
'
'inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
'inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
'inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
'inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
'inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
'inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
'inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
'inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
'Dim supplierCountry As String
'supplierCountry = Trim$(cIdentificationCode & "")
'If supplierCountry = "" Then supplierCountry = "SA"
'inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry
'inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
'inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")
'
'' ---------------------- Customer (Standard/Simplified) ----------------------
'Dim isStandard As Boolean
'isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")
'
'Dim isSimplified As Boolean
'isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
'               And (LCase$(mTableName) = "tbleinvoice" _
'                 Or LCase$(mTableName) = "propertyduebatchdetail" _
'                 Or LCase$(mTableName) = "debitandcreditnotification")
' Dim CUSTID As String
'If isStandard Or isSimplified Then
'
'    If Len(Trim$(Id700 & "")) > 0 Then
'        CUSTID = Trim$(Id700 & "")
'    Else
'        CUSTID = Trim$(Identificationid & "")
'    End If
'    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
'        CUSTID = "300000000000003": schemeID = "OTH"
'    End If
'
'    inv.CustomerParty.PartyIdentification.ID = CUSTID
'    If Len(Trim$(schemeID & "")) = 0 Then schemeID = "OTH"
'    inv.CustomerParty.PartyIdentification.schemeID = schemeID
'
'    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
'    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
'    cAddStreet = Trim$(AdditionalStreetName & "")
'    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
'    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
'    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
'    cSubent = Trim$(CountrySubentity & "")
'    cSubCity = Trim$(CitySubdivisionName & "")
'    cCountry = Trim$(IdentificationCode & "")
'    If Len(cCountry) = 0 Then cCountry = "SA"
'
'    inv.CustomerParty.PostalAddress.StreetName = cStreet
'    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
'    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
'    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
'    inv.CustomerParty.PostalAddress.CityName = cCity
'    inv.CustomerParty.PostalAddress.PostalZone = cPost
'    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
'    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
'    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry
'
'    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
'    If Len(Trim$(Id700 & "")) > 0 Then
'        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
'    Else
'        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
'    End If
'End If
'
'' ›Ì ﬁ”„ Customer - √÷› Â–« «· Õﬁﬁ:
'
'If isStandard Or isSimplified Then
'
'    If Len(Trim$(Id700 & "")) > 0 Then
'        CUSTID = Trim$(Id700 & "")
'        ' ===== —ﬁ„ 700 - «” Œœ„ MOM =====
'        schemeID = "MOM"
'    Else
'        CUSTID = Trim$(Identificationid & "")
'    End If
'
'    ' «· Õﬁﬁ Ê«· ’ÕÌÕ «· ·ﬁ«∆Ì
'    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
'        CUSTID = "300000000000003"
'        schemeID = "OTH"
'    Else
'        ' =====  ’ÕÌÕ  ·ﬁ«∆Ì »‰«¡ ⁄·Ï √Ê· —ﬁ„ =====
'        Dim firstChar As String
'        firstChar = left$(CUSTID, 1)
'
'        If firstChar = "7" And Len(CUSTID) = 10 Then
'            ' —ﬁ„ 700 - ÌÃ» √‰ ÌﬂÊ‰ MOM
'            schemeID = "MOM"
'        ElseIf firstChar = "1" And Len(CUSTID) = 10 Then
'            schemeID = "NAT"  ' „Ê«ÿ‰ ”⁄ÊœÌ
'        ElseIf firstChar = "2" And Len(CUSTID) = 10 Then
'            schemeID = "IQA"  ' „ﬁÌ„
'        ElseIf Len(CUSTID) = 15 Then
'            schemeID = "TIN"  ' —ﬁ„ ÷—Ì»Ì
'        ElseIf Len(CUSTID) = 10 Then
'            schemeID = "CRN"  ' ”Ã·  Ã«—Ì
'        Else
'            schemeID = "OTH"
'        End If
'    End If
'
'    inv.CustomerParty.PartyIdentification.ID = CUSTID
'    inv.CustomerParty.PartyIdentification.schemeID = schemeID
'End If
'' =========================================================
'' ============   ÿßŸÑÿ®ŸÜŸàÿØ + ÿ™ÿ¨ŸÖŸäÿπ ÿßŸÑÿÆÿµŸàŸÖÿßÿ™   ===============
'' =========================================================
'Dim StrSQL As String
'StrSQL = ""
'StrSQL = StrSQL & "SELECT "
'StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT,Trans_DiscountType,Trans_Discount, "
'StrSQL = StrSQL & "  TD.ShowQty, "
'StrSQL = StrSQL & "  TD.ItemDiscountType, "
'StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
'StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
'StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "
'
'If SystemOptions.PriceWithVAT Then
'    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
'    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
'Else
'    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
'    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
'End If
'
'StrSQL = StrSQL & "FROM dbo.Transactions AS T "
'StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
'StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
'StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID
'
'Dim rs As New ADODB.Recordset
'rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
'
'' *** ŸÖÿ™ÿ∫Ÿäÿ±ÿßÿ™ ŸÑŸÑÿ™ÿ¨ŸÖŸäÿπ ÿ®ÿØŸÇÿ© ÿπÿßŸÑŸäÿ© ***
'Dim SumExNoDisc As Double
'Dim SumTaxableExNoDisc As Double
'Dim SumExemptExNoDisc As Double
'Dim SumLineDiscEx As Double
'Dim mTrans_DiscountType As Integer
'Dim mTrans_Discount As Double
'
'SumExNoDisc = 0#
'SumTaxableExNoDisc = 0#
'SumExemptExNoDisc = 0#
'SumLineDiscEx = 0#
'
'If Not rs.EOF Then
'    mTrans_DiscountType = val(rs!Trans_DiscountType & "")
'    mTrans_Discount = val(rs!Trans_Discount & "")
'End If
'
'' =========================================================
'' ============   ŸÖÿπÿßŸÑÿ¨ÿ© ÿßŸÑÿ®ŸÜŸàÿØ (LOOP)   ==================
'' =========================================================
'Do Until rs.EOF
'    Dim lineEx As Double
'    Dim Rate As Double
'    Dim Qty As Double
'    Dim basePrice As Double
'    Dim discType As Long
'    Dim discPercent As Double
'    Dim discValue As Double
'    Dim discRaw As Double
'    Dim lineDiscEx As Double
'    Dim lineBaseForPercent As Double
'    Dim isItemExempt As Boolean
'
'    ' ŸÇŸäŸÖ ÿßŸÑÿ≥ÿ∑ÿ± ÿßŸÑÿ£ÿ≥ÿßÿ≥Ÿäÿ©
'    lineEx = val(rs!LineEx_NoDisc & "")
'    Rate = val(rs!vatRate & "")
'    Qty = val(rs!ShowQty & "")
'    basePrice = val(rs!BasePrice_Ex & "")
'    isItemExempt = (val(rs!ItemWithOutVAT & "") <> 0)
'
'    ' ÿ®ŸäÿßŸÜÿßÿ™ ÿßŸÑÿÆÿµŸÖ
'    discType = val(rs!ItemDiscountType & "")
'    discPercent = val(rs!ItemDiscount & "")
'    discValue = val(rs!ItemDiscount & "")
'    discRaw = 0
'
'    ' ===== ÿ≠ÿ≥ÿßÿ® ÿßŸÑÿÆÿµŸÖ ÿßŸÑŸÖÿÆÿ≤ŸÜ ÿπŸÑŸâ ÿßŸÑÿ≥ÿ∑ÿ± =====
'    Select Case discType
'        Case 3      ' ÿÆÿµŸÖ ŸÜÿ≥ÿ®ÿ© %
'            If SystemOptions.PriceWithVAT And Rate > 0 Then
'                lineBaseForPercent = lineEx * (1 + Rate / 100#)
'            Else
'                lineBaseForPercent = lineEx
'            End If
'            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) _
'                      + val(rs!TotalDiscountPerLine & "")
'
'        Case 2      ' ÿÆÿµŸÖ ŸÇŸäŸÖÿ©
'            discRaw = discValue + val(rs!TotalDiscountPerLine & "")
'
'        Case Else
'            discRaw = val(rs!TotalDiscountPerLine & "")
'    End Select
'
'    ' ===== ÿ™ÿ≠ŸàŸäŸÑ ÿßŸÑÿÆÿµŸÖ ÿ•ŸÑŸâ Ex-VAT =====
'    If SystemOptions.PriceWithVAT And Rate > 0 Then
'        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
'    Else
'        lineDiscEx = Round(discRaw, 6)
'    End If
'
'    ' ===== ÿ™ÿ¨ŸÖŸäÿπ ÿßŸÑÿ£ÿ≥ÿßÿ≥ÿßÿ™ ŸÇÿ®ŸÑ ÿßŸÑÿÆÿµŸÖ =====
'    SumExNoDisc = SumExNoDisc + lineEx
'
'    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
'        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
'    Else
'        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
'    End If
'
'    ' ===== ÿ™ÿ¨ŸÖŸäÿπ ÿÆÿµŸàŸÖÿßÿ™ ÿßŸÑÿ≥ÿ∑Ÿàÿ± =====
'    SumLineDiscEx = SumLineDiscEx + lineDiscEx
'
'    ' ===== ÿ•ŸÜÿ¥ÿßÿ° ÿ≥ÿ∑ÿ± UBL =====
'    Dim invline As Object
'    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
'
'    If docType = 0 Then
'        invline.Price.EncludingVat = SystemOptions.PriceWithVAT
'    Else
'        invline.Price.EncludingVat = False
'    End If
'
'    invline.Price.PriceAmount = Format(basePrice, "0.######")
'    invline.InvoiceQuantity = Format(Qty, "0.######")
'    invline.Item.Name = rs!ItemName & ""
'
'    ' =========================================================
'    ' ============   ŸÖÿπÿßŸÑÿ¨ÿ© ÿßŸÑÿ™ÿµŸÜŸäŸÅ ÿßŸÑÿ∂ÿ±Ÿäÿ®Ÿä (ÿßŸÑŸÖŸáŸÖÿ©!)   ======
'    ' =========================================================
'    Dim ratePct As Double
'    ratePct = IIf(Rate > 1, Rate, Rate * 100)  ' ÿ™ÿ∑ÿ®Ÿäÿπ ÿßŸÑŸÜÿ≥ÿ®ÿ©
'
'    If ratePct > 0 And Not isItemExempt Then
'        ' ===== ÿÆÿßÿ∂ÿπ ŸÑŸÑÿ∂ÿ±Ÿäÿ®ÿ© (S) =====
'        invline.Item.classifiedTaxCategory.ID = "S"
'        invline.Item.classifiedTaxCategory.Percent = ratePct
'        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
'        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = ratePct
'
'    Else
'        ' ===== ÿ∫Ÿäÿ± ÿÆÿßÿ∂ÿπ - ÿ™ÿ≠ÿØŸäÿØ ŸÜŸàÿπ ÿßŸÑÿ•ÿπŸÅÿßÿ° =====
'
'        ' 1) ÿ•ÿπŸÅÿßÿ° ÿπŸÇÿßÿ±Ÿä ÿ≥ŸÉŸÜŸä (docType = 6)
'        If docType = 6 And isItemExempt Then
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-30"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Real estate residential lease/sale"
'
'        ' 2) ÿ™ÿµÿØŸäÿ± ÿ≥ŸÑÿπ (ÿµŸÅÿ±Ÿä)
'        ElseIf InvoiceTypeCodename = "0100100" Then
'            invline.Item.classifiedTaxCategory.ID = "Z"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-32"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Export of goods"
'
'        ' 3) ÿÆÿØŸÖÿßÿ™ ŸÖÿßŸÑŸäÿ© ŸÖÿπŸÅÿßÿ© (ŸÖÿßÿØÿ© 29)
'        ElseIf InvoiceTypeCodename = "0100000" And docType <> 6 Then
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
'
'        ' 4) ÿÆÿßÿ±ÿ¨ ÿßŸÑŸÜÿ∑ÿßŸÇ (Out of Scope)
'        ElseIf InvoiceTypeCodename = "0100000" And docType = 6 Then
'            invline.Item.classifiedTaxCategory.ID = "O"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
'
'        ' 5) ŸÅÿßÿ™Ÿàÿ±ÿ© ŸÖÿ®ÿ≥ÿ∑ÿ© ÿ∫Ÿäÿ± ÿÆÿßÿ∂ÿπÿ©
'        ElseIf mId$(InvoiceTypeCodename & "", 1, 2) = "02" And ratePct = 0 Then
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
'
'        ' 6) ÿßŸÅÿ™ÿ±ÿßÿ∂Ÿä: ŸÖÿπŸÅŸâ ÿ®ÿØŸàŸÜ ÿ≥ÿ®ÿ® ÿÆÿßÿµ
'        Else
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
'        End If
'    End If
'
'    ' ===== ŸÖÿπÿßŸÑÿ¨ÿ© ÿÆÿßÿµÿ© ŸÑŸÄ DebitAndCreditNotification =====
'    If LCase$(mTableName) = "debitandcreditnotification" And ratePct = 0 Then
'        invline.Item.classifiedTaxCategory.ID = "O"
'        invline.Item.classifiedTaxCategory.Percent = 0
'        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
'        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
'    End If
'
'    ' ===== ŸÖÿπÿßŸÑÿ¨ÿ© ÿßŸÑÿπŸäŸÜÿßÿ™ ÿßŸÑŸÖÿ¨ÿßŸÜŸäÿ© (ÿßŸÑÿ≥ÿπÿ± = 0) =====
'    If Round(basePrice, 2) = 0 Then
'        invline.Item.classifiedTaxCategory.ID = "Z"
'        invline.Item.classifiedTaxCategory.Percent = 0
'        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
'        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-00"
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Free sample"
'    End If
'
'    inv.InvoiceLines.Add invline
'
'    rs.MoveNext
'Loop
'
'rs.Close
'
'' =========================================================
'' ==============  ÿÆÿµŸÖ ŸÖŸàÿ≠ŸëÿØ (ÿ≥ÿ∑Ÿàÿ± + ŸáŸäÿØÿ±)  ================
'' =========================================================
'Dim HeaderDiscEx As Double
'Dim TransDiscountRaw As Double
'Dim TransDiscountEx As Double
'
'HeaderDiscEx = Round(SumLineDiscEx, 2)
'TransDiscountRaw = val(Trans_Discount & "")
'
'If TransDiscountRaw <> 0 Then
'    If val(mTrans_DiscountType) = 1 Then
'        ' ÿÆÿµŸÖ ŸÜÿ≥ÿ®ÿ© %
'        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
'            Dim TotalWithVAT As Double
'            TotalWithVAT = SumExNoDisc * (1 + 0.15)
'            TransDiscountRaw = Round(TotalWithVAT * TransDiscountRaw / 100#, 6)
'            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
'        Else
'            TransDiscountRaw = Round(SumExNoDisc * TransDiscountRaw / 100#, 6)
'            TransDiscountEx = TransDiscountRaw
'        End If
'    Else
'        ' ÿÆÿµŸÖ ŸÇŸäŸÖÿ© ÿ´ÿßÿ®ÿ™ÿ©
'        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
'            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
'        Else
'            TransDiscountEx = TransDiscountRaw
'        End If
'    End If
'
'    HeaderDiscEx = HeaderDiscEx + TransDiscountEx
'End If
'
'If HeaderDiscEx < 0 Then HeaderDiscEx = 0
'
'' ÿ™Ÿàÿ≤Ÿäÿπ ÿØÿßÿÆŸÑŸä (ŸÜŸêÿ≥ÿ®Ÿä)
'Dim BaseTotal As Double: BaseTotal = SumExNoDisc
'Dim Allow_S As Double, Allow_E As Double
'
'If BaseTotal > 0 Then
'    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 2)
'    Allow_E = Round(HeaderDiscEx - Allow_S, 2)
'Else
'    Allow_S = 0
'    Allow_E = HeaderDiscEx
'End If
'
'' ===== ÿ≠ÿ≥ÿßÿ® ÿÆÿµŸÖ ÿßŸÑŸáŸäÿØÿ± ÿ¥ÿßŸÖŸÑÿßŸã ÿßŸÑÿ∂ÿ±Ÿäÿ®ÿ© =====
'Dim HeaderDiscVat As Double
'Dim Allow_S_Inc As Double
'Dim Allow_E_Inc As Double
'
'HeaderDiscVat = Round(Allow_S * 0.15, 2)
'Allow_S_Inc = Round(Allow_S, 2)
'Allow_E_Inc = Round(Allow_E, 2)
'
'' --------- ÿ£ÿ∂ŸÅ AllowanceCharge ŸÑŸÑŸáŸäÿØÿ± ---------
'If HeaderDiscEx > 0# Then
'
'    ' ÿ¨ÿ≤ÿ° ÿßŸÑÿÆÿµŸÖ ÿßŸÑŸàÿßŸÇÿπ ÿπŸÑŸâ ÿßŸÑÿ®ŸÜŸàÿØ ÿßŸÑÿÆÿßÿ∂ÿπÿ© S
'    If Allow_S_Inc > 0# Then
'        Dim allowanceS As Object
'        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'        allowanceS.ChargeIndicator = False
'        allowanceS.Amount = Format(Allow_S_Inc, "0.##")
'        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
'        allowanceS.taxCategory.ID = "S"
'        allowanceS.taxCategory.Percent = 15#
'        inv.allowanceCharges.Add allowanceS
'    End If
'
'    ' ÿ¨ÿ≤ÿ° ÿßŸÑÿÆÿµŸÖ ÿßŸÑŸàÿßŸÇÿπ ÿπŸÑŸâ ÿßŸÑÿ®ŸÜŸàÿØ ÿßŸÑŸÖÿπŸÅÿßÿ© E
'    If Allow_E_Inc > 0# Then
'        Dim allowanceE As Object
'        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'        allowanceE.ChargeIndicator = False
'        allowanceE.Amount = Format(Allow_E_Inc, "0.##")
'        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
'        allowanceE.taxCategory.ID = "E"
'        allowanceE.taxCategory.Percent = 0#
'        inv.allowanceCharges.Add allowanceE
'    End If
'
'End If
'
'
'' =========================================================
'' ====================   ÿßŸÑÿ•ÿ¨ŸÖÿßŸÑŸäÿßÿ™   =====================
'' =========================================================
'
'SumExNoDisc = Round(SumExNoDisc, 2)
'SumTaxableExNoDisc = Round(SumTaxableExNoDisc, 2)
'HeaderDiscEx = Round(HeaderDiscEx, 2)
'
'inv.LegalMonetaryTotal.LineExtensionAmount = Format(SumExNoDisc, "0.##")
'inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(HeaderDiscEx + HeaderDiscVat, "0.##")
'
'Dim TaxExclusiveAfterDisc As Double
'TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 2)
'If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0
'
'inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(TaxExclusiveAfterDisc, "0.##")
'
'' ===== ÿßÿ≠ÿ™ÿ≥ÿßÿ® ÿ∂ÿ±Ÿäÿ®ÿ© ÿßŸÑŸÇŸäŸÖÿ© ÿßŸÑŸÖÿ∂ÿßŸÅÿ© =====
'Dim TaxableAfterDisc As Double
'Dim TotalVat As Double
'
'TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 2)
'If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0
'
'Dim rawVat As Double
'rawVat = TaxableAfterDisc * 0.15
'TotalVat = Fix(rawVat * 100) / 100
'
'inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")
'
'Dim TaxInclusive As Double
'TaxInclusive = CDbl(Format(TaxExclusiveAfterDisc, "0.00")) + CDbl(Format(TotalVat, "0.00"))
'TaxInclusive = Round(TaxInclusive, 2)
'
'inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(TaxInclusive, "0.##")
'
'inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
'inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"
'
'Dim Prepaid As Double
'If val(InvoiceTypeCodeID) = 388 Then
'    Prepaid = val(PrepaidAmount & "")
'Else
'    Prepaid = 0
'End If
'inv.LegalMonetaryTotal.PrepaidAmount = Format(Round(Prepaid, 2), "0.##")
'
'Dim PayableAmountCalc As Double
'PayableAmountCalc = Round(TaxInclusive - Prepaid, 2)
'If PayableAmountCalc < 0 Then PayableAmountCalc = 0
'inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmountCalc, "0.##")
'
'' =========================================================
'' ==================   ÿßŸÑÿ™ŸàŸÇŸäÿπ ŸàÿßŸÑÿ•ÿ±ÿ≥ÿßŸÑ   =================
'' =========================================================
'Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
'Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey
'
'If SendingMode <> 0 Then
'    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
'    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
'End If
'
'Set res = ubl.GenerateInvoiceXML(inv, App.path)
'
'If res.IsValid = True Then
'    InvoiceHash = res.InvoiceHash
'    SingedXML = res.SingedXML
'    EncodedInvoice = res.EncodedInvoice
'    UUID = res.UUID
'    QRCode = res.QRCode
'    PIH = res.PIH
'    SingedXMLFileName = res.SingedXMLFileName
'
'    Dim apireqlogic As Object, tokenresponse As Object
'    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
'    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
'    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
'    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
'    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
'    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")
'
'    invrequestbody.invoice = res.EncodedInvoice
'    invrequestbody.InvoiceHash = res.InvoiceHash
'    invrequestbody.UUID = res.UUID
'
'    Select Case SendingMode
'        Case 0 ' DEV (Compliance)
'            apireqlogic.SetApiRequestLogic 0
'            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
'            If tokenresponse.errorMessage = "" Then
'                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
'                If invoicereportingmodel.IsSuccess = True Then
'                    If invoicereportingmodel.StatusCode = 202 Then
'                        WarningMessage = invoicereportingmodel.WarningMessage
'                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                    End If
'                    ErrorMessageS = invoicereportingmodel.ReportingStatus
'                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                    zatcaStatus = 1
'                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                    updatedatatable val(docType), mTableName, mFieldIDName
'                Else
'                    ErrorMessageS = invoicereportingmodel.errorMessage
'                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                End If
'            Else
'                ErrorMessageS = tokenresponse.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'
'        Case 1 ' SIMU (Reporting/Clearance)
'            apireqlogic.SetApiRequestLogic 1
'            apireqlogic.SetToSaveClearedXMLFile App.path, True
'            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
'            If invoiceclearingmodel.IsSuccess = True Then
'                If invoiceclearingmodel.StatusCode = 202 Then
'                    WarningMessage = invoiceclearingmodel.WarningMessage
'                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                End If
'                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                QRCode = invoiceclearingmodel.QRCode
'                zatcaStatus = 1
'                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                updatedatatable val(docType), mTableName, mFieldIDName
'            Else
'                ErrorMessageS = invoiceclearingmodel.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'
'        Case 2 ' PROD (Clearance)
'            apireqlogic.SetApiRequestLogic 2
'            apireqlogic.SetToSaveClearedXMLFile App.path, True
'            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
'            If invoiceclearingmodel.IsSuccess = True Then
'                If invoiceclearingmodel.StatusCode = 202 Then
'                    WarningMessage = invoiceclearingmodel.WarningMessage
'                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                End If
'                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                QRCode = invoiceclearingmodel.QRCode
'                zatcaStatus = 1
'                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                updatedatatable val(docType), mTableName, mFieldIDName
'            Else
'                ErrorMessageS = invoiceclearingmodel.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'    End Select
'
'    GenerateInvoice_ZatcaV3 = True
'Else
'    ErrorMessageS = res.errorMessage
'    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'    GenerateInvoice_ZatcaV3 = False
'End If
'
'End Function
'
'
' =====================================================
' GenerateInvoice_ZatcaV3 - «·‰”Œ… «·ﬂ«„·… «·„’ÕÕ…
'  «—ÌŒ «· ÕœÌÀ: 2025-12-29
' =====================================================
' «·≈’·«Õ«  «·„÷„‰…:
' 1. BR-CL-14:  ‰ŸÌ› ﬂÊœ «·œÊ·…
' 2. BR-KSA-CL-04: TaxExemptionReasonCode ›Ì ClassifiedTaxCategory
' 3. BR-KSA-F-10:  ’ÕÌÕ schemeID (—ﬁ„ 700 = MOM)
' 4. BR-KSA-F-13: «· Õﬁﬁ „‰ Buyer ID ··›« Ê—… «·÷—Ì»Ì…
' 5. DocumentCurrencyCode Ê TaxCurrencyCode
' 6. Ã„Ì⁄ Õ«·«  «·≈⁄›«¡ «·÷—Ì»Ì
' =====================================================

'Public Function GenerateInvoice_ZatcaV3( _
'    Optional ByVal docType As Integer = 0, _
'    Optional ByVal mTableName As String = "Transactions", _
'    Optional ByVal mFieldIDName As String = "Transaction_ID", _
'    Optional ByVal ActivityTypeIdInv As Integer = 0, _
'    Optional ByVal branch_idInv As Integer = 0 _
') As Boolean
'
'On Error Resume Next
'
'    ' === ≈–« «·√”⁄«— ‘«„·… «·÷—Ì»…° «” Œœ„ «·œ«·… «·ﬁœÌ„… ===
'    If SystemOptions.PriceWithVAT Then
'        Call generateInvoice(docType, mTableName, mFieldIDName, ActivityTypeIdInv, branch_idInv)
'        GenerateInvoice_ZatcaV3 = True
'        Exit Function
'    End If
'
'' =====================  Õ÷Ì— «·‹ UBL/Invoice =====================
'Dim ubl As Object, inv As Object, res As Object
'Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
'Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
'Set res = CreateObject("ZatcaUBLXMLGenerator.Result")
'
'' ---- « ’«· ﬁ«⁄œ… «·»Ì«‰«  ----
'If SystemOptions.CanUploadZakatOpt Then
'    ConectionFirst
'    Set cn2 = POSConnection
'Else
'    Set cn2 = Cn
'End If
'
'' ---- »Ì«‰«  «·‘—ﬂ…/«·›—⁄ ----
'If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
'    fillmycompanydata ActivityTypeIdInv, branch_idInv
'Else
'    fillmycompanydata
'End If
'
'' ---------------------- BillingReference (CN/DN only) ----------------------
'If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
'    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
'        Dim InvoiceDocumentReference As Object
'        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
'        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
'        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
'    End If
'End If
'
'' ---------------------- ‰Ê⁄ «·›« Ê—… ----------------------
'inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
'inv.InvoiceTypeCode.Name = InvoiceTypeCodename
'
'' ---------------------- «·⁄„·… (≈’·«Õ „Â„!) ----------------------
'If DocumentCurrencyCode = "" Or DocumentCurrencyCode = "SR" Then DocumentCurrencyCode = "SAR"
'If TaxCurrencyCode = "" Or TaxCurrencyCode = "SR" Then TaxCurrencyCode = "SAR"
'inv.DocumentCurrencyCode = DocumentCurrencyCode
'inv.TaxCurrencyCode = TaxCurrencyCode
'
'' ---------------------- ICV / UUID / PIH ----------------------
'inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
'
'Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
'icvOnlyDigits = ""
'icvStr = Trim$(uuidCounter & "")
'If icvStr = "" Then icvStr = "1"
'For i = 1 To Len(icvStr)
'    Ch = mId$(icvStr, i, 1)
'    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
'Next i
'If icvOnlyDigits = "" Then icvOnlyDigits = "1"
'inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits
'
'inv.ID = ID
'
'' ---------------------- Issue Date & Time ----------------------
'If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
'inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")
'
'If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
'inv.IssueTime = Format(IssueTim, "HH:mm:ss")
'
'' ---------------------- Delivery Dates ----------------------
'If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
'If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
'inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
'If val(InvoiceTypeCodeID) = 388 Then
'    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
'End If
'
'' ---------------------- Payment Means ----------------------
'Dim paymentmean As Object
'Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
'paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
'If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
'    paymentmean.InstructionNote = "Returned items"
'End If
'If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
'    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
'End If
'If Len(Trim$(paymentnote & "")) > 0 Then
'    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
'End If
'inv.PaymentMeans.Add paymentmean
'
'' ---------------------- Supplier ----------------------
'inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
'If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
'inv.SupplierParty.PartyIdentification.schemeID = cschemeID
'
'inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
'inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
'inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
'inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
'inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
'inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
'inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
'inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
'
'' === ≈’·«Õ ﬂÊœ œÊ·… «·„Ê—œ ===
'Dim supplierCountry As String
'supplierCountry = CleanCountryCode(cIdentificationCode)
'inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry
'
'inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
'inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")
'
'' =========================================================
'' ============   »Ì«‰«  «·⁄„Ì· (Customer)   ==============
'' =========================================================
'Dim isStandard As Boolean
'Dim isSimplified As Boolean
'
'isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")
'isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
'               And (LCase$(mTableName) = "tbleinvoice" _
'                 Or LCase$(mTableName) = "propertyduebatchdetail" _
'                 Or LCase$(mTableName) = "debitandcreditnotification")
'
'If isStandard Or isSimplified Then
'    Dim CUSTID As String
'    Dim custSchemeID As String
'
'    ' ===  ÕœÌœ —ﬁ„ «·⁄„Ì· ===
'    If Len(Trim$(Id700 & "")) > 0 Then
'        CUSTID = Trim$(Id700 & "")
'        custSchemeID = "MOM"  ' —ﬁ„ 700 œ«∆„« MOM
'    Else
'        CUSTID = Trim$(Identificationid & "")
'        custSchemeID = Trim$(schemeID & "")
'    End If
'
'    ' === «· Õﬁﬁ „‰ «·—ﬁ„ «·«› —«÷Ì ··›« Ê—… «·÷—Ì»Ì… (BR-KSA-F-13) ===
''    If isStandard Then
''        If CUSTID = "300000000000003" Or CUSTID = "" Or CUSTID = "0" Or UCase$(CUSTID) = "N/A" Then
''            ' «·›« Ê—… «·÷—Ì»Ì…   ÿ·» —ﬁ„  ⁄—Ì› ÕﬁÌﬁÌ
''            ' «” Œœ„ «·—ﬁ„ «·÷—Ì»Ì ≈–« ﬂ«‰ „ Ê›—«
''            If Len(Trim$(CompanyID & "")) = 15 And IsNumeric(CompanyID) Then
''                CUSTID = Trim$(CompanyID)
''                custSchemeID = "TIN"
''            Else
''                ' ·« ÌÊÃœ —ﬁ„  ⁄—Ì› ’«·Õ - Œÿ√!
''                ErrorMessageS = "Error: Standard invoice requires valid Buyer ID (CR/TIN/NAT). Customer ID is missing or invalid."
''                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
''                GenerateInvoice_ZatcaV3 = False
''                Exit Function
''            End If
''        End If
''    Else
''        ' ›« Ê—… „»”ÿ… - «·—ﬁ„ «·«› —«÷Ì „ﬁ»Ê·
''        If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
''            CUSTID = "300000000000003"
''            custSchemeID = "OTH"
''        End If
''    End If
''
'
'    ' === «· Õﬁﬁ „‰ »Ì«‰«  «·⁄„Ì· ===
'If isStandard Then
'    If CUSTID = "300000000000003" Or CUSTID = "" Or CUSTID = "0" Or UCase$(CUSTID) = "N/A" Then
'        ' Õ«Ê· «” Œœ«„ «·—ﬁ„ «·÷—Ì»Ì √Ê·«
'        If Len(Trim$(CompanyID & "")) = 15 And IsNumeric(CompanyID) Then
'            CUSTID = Trim$(CompanyID)
'            custSchemeID = "TIN"
'            custSchemeID = "OTH"
'        Else
'            ' ===== ⁄„Ì· ‰ﬁœÌ - «” Œœ„ «·—ﬁ„ «·«› —«÷Ì =====
'            ' ZATCA  ﬁ»· «·—ﬁ„ «·«› —«÷Ì ··⁄„·«¡ «·‰ﬁœÌÌ‰
'            CUSTID = "300000000000003"
'            custSchemeID = "OTH"
'            custSchemeID = "TIN"
'        End If
'    End If
'Else
'    ' ›« Ê—… „»”ÿ…
'    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
'        CUSTID = "300000000000003"
'        custSchemeID = "OTH"
'    End If
'End If
'
'    ' ===  ’ÕÌÕ schemeID  ·ﬁ«∆Ì« (BR-KSA-F-10) ===
'    custSchemeID = GetCorrectSchemeID(CUSTID, custSchemeID)
'
'    inv.CustomerParty.PartyIdentification.ID = CUSTID
'    inv.CustomerParty.PartyIdentification.schemeID = custSchemeID
'
'    ' === ⁄‰Ê«‰ «·⁄„Ì· ===
'    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
'    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
'    cAddStreet = Trim$(AdditionalStreetName & "")
'    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
'    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
'    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
'    cSubent = Trim$(CountrySubentity & "")
'    cSubCity = Trim$(CitySubdivisionName & "")
'
'    ' === ≈’·«Õ ﬂÊœ œÊ·… «·⁄„Ì· (BR-CL-14) ===
'    cCountry = CleanCountryCode(IdentificationCode)
'
'    inv.CustomerParty.PostalAddress.StreetName = cStreet
'    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
'    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
'    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
'    inv.CustomerParty.PostalAddress.CityName = cCity
'    inv.CustomerParty.PostalAddress.PostalZone = cPost
'    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
'    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
'    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry
'
'    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
'
'    ' === «·—ﬁ„ «·÷—Ì»Ì ··⁄„Ì· ===
'    If Len(Trim$(Id700 & "")) > 0 Then
'        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
'    Else
'        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
'    End If
'End If
'
'' =========================================================
'' ============   «·»‰Êœ +  Ã„Ì⁄ «·Œ’Ê„«    ===============
'' =========================================================
'Dim StrSQL As String
'StrSQL = ""
'StrSQL = StrSQL & "SELECT "
'StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
'StrSQL = StrSQL & "  T.Trans_DiscountType, T.Trans_Discount, "
'StrSQL = StrSQL & "  TD.ShowQty, "
'StrSQL = StrSQL & "  TD.ItemDiscountType, "
'StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
'StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
'StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "
'
'If SystemOptions.PriceWithVAT Then
'    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
'    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
'Else
'    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
'    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
'End If
'
'StrSQL = StrSQL & "FROM dbo.Transactions AS T "
'StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
'StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
'StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID
'
'Dim rs As New ADODB.Recordset
'rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
'
'' *** „ €Ì—«  ·· Ã„Ì⁄ ***
'Dim SumExNoDisc As Double
'Dim SumTaxableExNoDisc As Double
'Dim SumExemptExNoDisc As Double
'Dim SumLineDiscEx As Double
'Dim mTrans_DiscountType As Integer
'Dim mTrans_Discount As Double
'
'SumExNoDisc = 0#
'SumTaxableExNoDisc = 0#
'SumExemptExNoDisc = 0#
'SumLineDiscEx = 0#
'
'If Not rs.EOF Then
'    mTrans_DiscountType = val(rs!Trans_DiscountType & "")
'    mTrans_Discount = val(rs!Trans_Discount & "")
'End If
'
'' =========================================================
'' ============   „⁄«·Ã… «·»‰Êœ (LOOP)   ==================
'' =========================================================
'Do Until rs.EOF
'    Dim lineEx As Double
'    Dim Rate As Double
'    Dim Qty As Double
'    Dim basePrice As Double
'    Dim discType As Long
'    Dim discPercent As Double
'    Dim discValue As Double
'    Dim discRaw As Double
'    Dim lineDiscEx As Double
'    Dim lineBaseForPercent As Double
'    Dim isItemExempt As Boolean
'    Dim ratePct As Double
'
'    ' ﬁÌ„ «·”ÿ— «·√”«”Ì…
'    lineEx = val(rs!LineEx_NoDisc & "")
'    Rate = val(rs!vatRate & "")
'    Qty = val(rs!ShowQty & "")
'    basePrice = val(rs!BasePrice_Ex & "")
'    isItemExempt = (val(rs!ItemWithOutVAT & "") <> 0)
'
'    '  ÿ»Ì⁄ «·‰”»…: ·Ê 0.15 ‰Œ·ÌÂ« 15
'    ratePct = IIf(Rate > 1, Rate, Rate * 100)
'
'    ' »Ì«‰«  «·Œ’„
'    discType = val(rs!ItemDiscountType & "")
'    discPercent = val(rs!ItemDiscount & "")
'    discValue = val(rs!ItemDiscount & "")
'    discRaw = 0
'
'    ' ===== Õ”«» «·Œ’„ «·„Œ“‰ ⁄·Ï «·”ÿ— =====
'    Select Case discType
'        Case 3      ' Œ’„ ‰”»… %
'            If SystemOptions.PriceWithVAT And Rate > 0 Then
'                lineBaseForPercent = lineEx * (1 + Rate / 100#)
'            Else
'                lineBaseForPercent = lineEx
'            End If
'            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) _
'                      + val(rs!TotalDiscountPerLine & "")
'
'        Case 2      ' Œ’„ ﬁÌ„…
'            discRaw = discValue + val(rs!TotalDiscountPerLine & "")
'
'        Case Else
'            discRaw = val(rs!TotalDiscountPerLine & "")
'    End Select
'
'    ' =====  ÕÊÌ· «·Œ’„ ≈·Ï Ex-VAT =====
'    If SystemOptions.PriceWithVAT And Rate > 0 Then
'        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
'    Else
'        lineDiscEx = Round(discRaw, 6)
'    End If
'
'    ' =====  Ã„Ì⁄ «·√”«”«  ﬁ»· «·Œ’„ =====
'    SumExNoDisc = SumExNoDisc + lineEx
'
'    If ratePct > 0 And Not isItemExempt Then
'        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
'    Else
'        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
'    End If
'
'    ' =====  Ã„Ì⁄ Œ’Ê„«  «·”ÿÊ— =====
'    SumLineDiscEx = SumLineDiscEx + lineDiscEx
'
'    ' ===== ≈‰‘«¡ ”ÿ— UBL =====
'    Dim invline As Object
'    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
'
'    If docType = 0 Then
'        invline.Price.EncludingVat = SystemOptions.PriceWithVAT
'    Else
'        invline.Price.EncludingVat = False
'    End If
'
'    invline.Price.PriceAmount = Format(basePrice, "0.######")
'    invline.InvoiceQuantity = Format(Qty, "0.######")
'    invline.Item.Name = rs!ItemName & ""
'
'    ' =========================================================
'    ' ===   „⁄«·Ã… «· ’‰Ì› «·÷—Ì»Ì (BR-KSA-CL-04 FIX)   ======
'    ' =========================================================
'
'    ' === «·⁄Ì‰«  «·„Ã«‰Ì… («·”⁄— = 0) - √Ê·ÊÌ… ⁄·Ì« ===
'    If Round(basePrice, 2) = 0 Then
'        ' Zero-rated - ⁄Ì‰… „Ã«‰Ì…
'        invline.Item.classifiedTaxCategory.ID = "Z"
'        invline.Item.classifiedTaxCategory.Percent = 0
'        invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-00"
'        invline.Item.classifiedTaxCategory.TaxExemptionReason = "Free sample"
'
'        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
'        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-00"
'        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Free sample"
'
'    ElseIf ratePct > 0 And Not isItemExempt Then
'        ' ===== Œ«÷⁄ ··÷—Ì»… (S) =====
'        invline.Item.classifiedTaxCategory.ID = "S"
'        invline.Item.classifiedTaxCategory.Percent = ratePct
'        ' S ·« ÌÕ «Ã TaxExemptionReasonCode
'
'        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
'        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = ratePct
'
'    Else
'        ' ===== €Ì— Œ«÷⁄ -  ÕœÌœ ‰Ê⁄ «·≈⁄›«¡ =====
'
'        ' 1) ≈⁄›«¡ ⁄ﬁ«—Ì ”ﬂ‰Ì (docType = 6)
'        If docType = 6 And isItemExempt Then
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-30"
'            invline.Item.classifiedTaxCategory.TaxExemptionReason = "Real estate residential lease/sale"
'
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-30"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Real estate residential lease/sale"
'
'        ' 2)  ’œÌ— ”·⁄ (’›—Ì) - InvoiceTypeCodename = "0100100"
'        ElseIf InvoiceTypeCodename = "0100100" Then
'            invline.Item.classifiedTaxCategory.ID = "Z"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-32"
'            invline.Item.classifiedTaxCategory.TaxExemptionReason = "Export of goods"
'
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-32"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Export of goods"
'
'        ' 3) Œ«—Ã «·‰ÿ«ﬁ (Out of Scope) - docType = 6 „⁄ InvoiceTypeCodename = "0100000"
'        ElseIf InvoiceTypeCodename = "0100000" And docType = 6 Then
'            invline.Item.classifiedTaxCategory.ID = "O"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'            invline.Item.classifiedTaxCategory.TaxExemptionReason = "Out of scope of VAT"
'
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
'
'        ' 4) DebitAndCreditNotification »œÊ‰ ÷—Ì»…
'        ElseIf LCase$(mTableName) = "debitandcreditnotification" And ratePct = 0 Then
'            invline.Item.classifiedTaxCategory.ID = "O"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'            invline.Item.classifiedTaxCategory.TaxExemptionReason = "Out of scope of VAT"
'
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-OOS"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of scope of VAT"
'
'        ' 5) √œÊÌ… Ê„⁄œ«  ÿ»Ì… (’›—Ì) - Ì„ﬂ‰  ›⁄Ì·Â« Õ”» ‰Ê⁄ «·‰‘«ÿ
'        ' ElseIf IsMedicalItem(rs!ItemID) Then
'        '     invline.Item.classifiedTaxCategory.ID = "Z"
'        '     invline.Item.classifiedTaxCategory.Percent = 0
'        '     invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-35"
'        '     invline.Item.classifiedTaxCategory.TaxExemptionReason = "Medicines and medical equipment"
'
'        ' 6) «› —«÷Ì: Œœ„«  „«·Ì… „⁄›«… («·„«œ… 29)
'        Else
'            invline.Item.classifiedTaxCategory.ID = "E"
'            invline.Item.classifiedTaxCategory.Percent = 0
'            invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
'            invline.Item.classifiedTaxCategory.TaxExemptionReason = "Financial services specified in Article 29"
'
'            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
'            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29"
'            invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29"
'        End If
'    End If
'
'    inv.InvoiceLines.Add invline
'
'    rs.MoveNext
'Loop
'
'rs.Close
'
'' =========================================================
'' ==============  Œ’„ „ÊÕ¯œ (”ÿÊ— + ÂÌœ—)  ================
'' =========================================================
'Dim HeaderDiscEx As Double
'Dim TransDiscountRaw As Double
'Dim TransDiscountEx As Double
'
'HeaderDiscEx = Round(SumLineDiscEx, 2)
'TransDiscountRaw = val(Trans_Discount & "")
'
'If TransDiscountRaw <> 0 Then
'    If val(mTrans_DiscountType) = 1 Then
'        ' Œ’„ ‰”»… %
'        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
'            Dim TotalWithVAT As Double
'            TotalWithVAT = SumExNoDisc * (1 + 0.15)
'            TransDiscountRaw = Round(TotalWithVAT * TransDiscountRaw / 100#, 6)
'            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
'        Else
'            TransDiscountRaw = Round(SumExNoDisc * TransDiscountRaw / 100#, 6)
'            TransDiscountEx = TransDiscountRaw
'        End If
'    Else
'        ' Œ’„ ﬁÌ„… À«» …
'        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
'            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
'        Else
'            TransDiscountEx = TransDiscountRaw
'        End If
'    End If
'
'    HeaderDiscEx = HeaderDiscEx + TransDiscountEx
'End If
'
'If HeaderDiscEx < 0 Then HeaderDiscEx = 0
'
''  Ê“Ì⁄ œ«Œ·Ì (‰ˆ”»Ì)
'Dim BaseTotal As Double: BaseTotal = SumExNoDisc
'Dim Allow_S As Double, Allow_E As Double
'
'If BaseTotal > 0 Then
'    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 2)
'    Allow_E = Round(HeaderDiscEx - Allow_S, 2)
'Else
'    Allow_S = 0
'    Allow_E = HeaderDiscEx
'End If
'
'' ===== Õ”«» Œ’„ «·ÂÌœ— ‘«„·« «·÷—Ì»… =====
'Dim HeaderDiscVat As Double
'Dim Allow_S_Inc As Double
'Dim Allow_E_Inc As Double
'
'HeaderDiscVat = Round(Allow_S * 0.15, 2)
'Allow_S_Inc = Round(Allow_S, 2)
'Allow_E_Inc = Round(Allow_E, 2)
'
'' --------- √÷› AllowanceCharge ··ÂÌœ— ---------
'If HeaderDiscEx > 0# Then
'
'    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S
'    If Allow_S_Inc > 0# Then
'        Dim allowanceS As Object
'        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'        allowanceS.ChargeIndicator = False
'        allowanceS.Amount = Format(Allow_S_Inc, "0.##")
'        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
'        allowanceS.taxCategory.ID = "S"
'        allowanceS.taxCategory.Percent = 15#
'        inv.allowanceCharges.Add allowanceS
'    End If
'
'    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·„⁄›«… E
'    If Allow_E_Inc > 0# Then
'        Dim allowanceE As Object
'        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'        allowanceE.ChargeIndicator = False
'        allowanceE.Amount = Format(Allow_E_Inc, "0.##")
'        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
'        allowanceE.taxCategory.ID = "E"
'        allowanceE.taxCategory.Percent = 0#
'        inv.allowanceCharges.Add allowanceE
'    End If
'
'End If
'
'
'' =========================================================
'' ====================   «·≈Ã„«·Ì«    =====================
'' =========================================================
'
'SumExNoDisc = Round(SumExNoDisc, 2)
'SumTaxableExNoDisc = Round(SumTaxableExNoDisc, 2)
'HeaderDiscEx = Round(HeaderDiscEx, 2)
'
'inv.LegalMonetaryTotal.LineExtensionAmount = Format(SumExNoDisc, "0.##")
'inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(HeaderDiscEx + HeaderDiscVat, "0.##")
'
'Dim TaxExclusiveAfterDisc As Double
'TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 2)
'If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0
'
'inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(TaxExclusiveAfterDisc, "0.##")
'
'' ===== «Õ ”«» ÷—Ì»… «·ﬁÌ„… «·„÷«›… =====
'Dim TaxableAfterDisc As Double
'Dim TotalVat As Double
'
'TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 2)
'If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0
'
'Dim rawVat As Double
'rawVat = TaxableAfterDisc * 0.15
'TotalVat = Fix(rawVat * 100) / 100
'
'inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")
'
'Dim TaxInclusive As Double
'TaxInclusive = CDbl(Format(TaxExclusiveAfterDisc, "0.00")) + CDbl(Format(TotalVat, "0.00"))
'TaxInclusive = Round(TaxInclusive, 2)
'
'inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(TaxInclusive, "0.##")
'
'inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
'inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"
'
'Dim Prepaid As Double
'If val(InvoiceTypeCodeID) = 388 Then
'    Prepaid = val(PrepaidAmount & "")
'Else
'    Prepaid = 0
'End If
'inv.LegalMonetaryTotal.PrepaidAmount = Format(Round(Prepaid, 2), "0.##")
'
'Dim PayableAmountCalc As Double
'PayableAmountCalc = Round(TaxInclusive - Prepaid, 2)
'If PayableAmountCalc < 0 Then PayableAmountCalc = 0
'inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmountCalc, "0.##")
'
'' =========================================================
'' ==================   «· ÊﬁÌ⁄ Ê«·≈—”«·   =================
'' =========================================================
'Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
'Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey
'
'If SendingMode <> 0 Then
'    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
'    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
'End If
'
'Set res = ubl.GenerateInvoiceXML(inv, App.path)
'
'If res.IsValid = True Then
'    InvoiceHash = res.InvoiceHash
'    SingedXML = res.SingedXML
'    EncodedInvoice = res.EncodedInvoice
'    UUID = res.UUID
'    QRCode = res.QRCode
'    PIH = res.PIH
'    SingedXMLFileName = res.SingedXMLFileName
'
'    Dim apireqlogic As Object, tokenresponse As Object
'    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
'    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
'    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
'    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
'    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
'    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")
'
'    invrequestbody.invoice = res.EncodedInvoice
'    invrequestbody.InvoiceHash = res.InvoiceHash
'    invrequestbody.UUID = res.UUID
'
'    Select Case SendingMode
'        Case 0 ' DEV (Compliance)
'            apireqlogic.SetApiRequestLogic 0
'            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
'            If tokenresponse.errorMessage = "" Then
'                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
'                If invoicereportingmodel.IsSuccess = True Then
'                    If invoicereportingmodel.StatusCode = 202 Then
'                        WarningMessage = invoicereportingmodel.WarningMessage
'                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                    End If
'                    ErrorMessageS = invoicereportingmodel.ReportingStatus
'                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                    zatcaStatus = 1
'                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                    updatedatatable val(docType), mTableName, mFieldIDName
'                Else
'                    ErrorMessageS = invoicereportingmodel.errorMessage
'                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                End If
'            Else
'                ErrorMessageS = tokenresponse.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'
'        Case 1 ' SIMU (Reporting/Clearance)
'            apireqlogic.SetApiRequestLogic 1
'            apireqlogic.SetToSaveClearedXMLFile App.path, True
'            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
'            If invoiceclearingmodel.IsSuccess = True Then
'                If invoiceclearingmodel.StatusCode = 202 Then
'                    WarningMessage = invoiceclearingmodel.WarningMessage
'                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                End If
'                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                QRCode = invoiceclearingmodel.QRCode
'                zatcaStatus = 1
'                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                updatedatatable val(docType), mTableName, mFieldIDName
'            Else
'                ErrorMessageS = invoiceclearingmodel.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'
'        Case 2 ' PROD (Clearance)
'            apireqlogic.SetApiRequestLogic 2
'            apireqlogic.SetToSaveClearedXMLFile App.path, True
'            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
'            If invoiceclearingmodel.IsSuccess = True Then
'                If invoiceclearingmodel.StatusCode = 202 Then
'                    WarningMessage = invoiceclearingmodel.WarningMessage
'                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
'                End If
'                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'                QRCode = invoiceclearingmodel.QRCode
'                zatcaStatus = 1
'                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
'                updatedatatable val(docType), mTableName, mFieldIDName
'            Else
'                ErrorMessageS = invoiceclearingmodel.errorMessage
'                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'            End If
'    End Select
'
'    GenerateInvoice_ZatcaV3 = True
'Else
'    ErrorMessageS = res.errorMessage
'    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
'    GenerateInvoice_ZatcaV3 = False
'End If
'
'End Function
'
'
' =====================================================
' œ«·… „”«⁄œ…:  ‰ŸÌ› ﬂÊœ «·œÊ·… (BR-CL-14 Fix)
' =====================================================
'Private Function CleanCountryCode(ByVal rawCode As String) As String
'    Dim cleanCode As String
'    Dim j As Integer
'    Dim c As String
'
'    cleanCode = ""
'    rawCode = Trim$(rawCode & "")
'
'    ' «” Œ—«Ã «·√Õ—› «·≈‰Ã·Ì“Ì… ›ﬁÿ A-Z
'    For j = 1 To Len(rawCode)
'        c = mId$(rawCode, j, 1)
'        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
'            cleanCode = cleanCode & UCase$(c)
'        End If
'    Next j
'
'    ' «· Õﬁﬁ „‰ «·’Õ… - ÌÃ» √‰ ÌﬂÊ‰ Õ—›Ì‰
'    If Len(cleanCode) = 2 Then
'        CleanCountryCode = cleanCode
'    Else
'        CleanCountryCode = "SA"  ' «› —«÷Ì
'    End If
'End Function
'
'
' =====================================================
' œ«·… „”«⁄œ…:  ’ÕÌÕ schemeID (BR-KSA-F-10 Fix)
' =====================================================
'Private Function GetCorrectSchemeID(ByVal BuyerID As String, ByVal currentSchemeID As String) As String
'    BuyerID = Trim$(BuyerID & "")
'    currentSchemeID = Trim$(UCase$(currentSchemeID & ""))
'
'    If Len(BuyerID) = 0 Then
'        GetCorrectSchemeID = "OTH"
'        Exit Function
'    End If
'
'    Dim firstChar As String
'    firstChar = left$(BuyerID, 1)
'
'    Dim idLength As Integer
'    idLength = Len(BuyerID)
'
'    ' ===== «· ’ÕÌÕ «· ·ﬁ«∆Ì =====
'
'    ' 15 —ﬁ„ = TIN (—ﬁ„ ÷—Ì»Ì)
'    If idLength = 15 And IsNumeric(BuyerID) Then
'        GetCorrectSchemeID = "TIN"
'        Exit Function
'    End If
'
'    ' 10 √—ﬁ«„
'    If idLength = 10 And IsNumeric(BuyerID) Then
'        Select Case firstChar
'            Case "1"
'                ' Ì»œ√ »‹ 1 - NAT √Ê CRN „ﬁ»Ê·
'                If currentSchemeID = "NAT" Or currentSchemeID = "CRN" Then
'                    GetCorrectSchemeID = currentSchemeID
'                Else
'                    GetCorrectSchemeID = "CRN"
'                End If
'
'            Case "2"
'                ' Ì»œ√ »‹ 2 - „ﬁÌ„ (IQA √Ê NAT)
'                If currentSchemeID = "NAT" Or currentSchemeID = "IQA" Then
'                    GetCorrectSchemeID = currentSchemeID
'                Else
'                    GetCorrectSchemeID = "IQA"
'                End If
'
'            Case "7"
'                ' ===== —ﬁ„ 700 / Œ·ÌÃÌ - ÌÃ» MOM Ê·Ì” NAT! =====
'                If currentSchemeID = "NAT" Then
'                    GetCorrectSchemeID = "MOM"  '  ’ÕÌÕ!
'                ElseIf currentSchemeID = "MOM" Or currentSchemeID = "GCC" Then
'                    GetCorrectSchemeID = currentSchemeID
'                Else
'                    GetCorrectSchemeID = "MOM"
'                End If
'
'            Case "3", "4", "5", "6", "8", "9", "0"
'                ' √—ﬁ«„ √Œ—Ï - ”Ã·  Ã«—Ì
'                If currentSchemeID = "NAT" Then
'                    GetCorrectSchemeID = "CRN"  ' NAT ·« Ì»œ√ »Â–Â «·√—ﬁ«„
'                ElseIf currentSchemeID = "CRN" Then
'                    GetCorrectSchemeID = "CRN"
'                Else
'                    GetCorrectSchemeID = "CRN"
'                End If
'        End Select
'        Exit Function
'    End If
'
'    ' €Ì— –·ﬂ
'    If currentSchemeID = "" Then
'        GetCorrectSchemeID = "OTH"
'    Else
'        GetCorrectSchemeID = currentSchemeID
'    End If
'End Function
'
Public Function GenerateInvoice_ZatcaV3DD( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "Transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0 _
) As Boolean
On Error Resume Next
   If SystemOptions.PriceWithVAT Then
        ' «” œ⁄ˆ «·›‰ﬂ‘‰ «·ﬁœÌ„… ﬂ„«  ⁄Ê¯œ 
        Call generateInvoice(docType, mTableName, mFieldIDName, ActivityTypeIdInv, branch_idInv)

        ' ·Ê  Õ»  —Ã⁄ True ﬂ⁄·«„… ‰Ã«Õ
        GenerateInvoice_ZatcaV3DD = True
        Exit Function
    End If
' =====================  Õ÷Ì— «·‹ UBL/Invoice =====================
Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

' ---- « ’«· (ﬂ„« ÂÊ »œÊ‰  ⁄œÌ·) ----
If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If

' ---- »Ì«‰«  «·‘—ﬂ…/«·›—⁄ ----
If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

' ---------------------- BillingReference (CN/DN only) ----------------------
If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If
End If

' ---------------------- ‰Ê⁄ «·›« Ê—… ----------------------
inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
inv.InvoiceTypeCode.Name = InvoiceTypeCodename

' ---------------------- ICV / UUID / PIH ----------------------
inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)

Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
icvOnlyDigits = ""
icvStr = Trim$(uuidCounter & "")
If icvStr = "" Then icvStr = "1"
For i = 1 To Len(icvStr)
    Ch = mId$(icvStr, i, 1)
    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
Next i
If icvOnlyDigits = "" Then icvOnlyDigits = "1"
inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits

inv.ID = ID

' ---------------------- Issue Date & Time ----------------------
If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")

If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
' „Â„: mm = œﬁ«∆ﬁ / ss = ÀÊ«‰Ì
inv.IssueTime = Format(IssueTim, "HH:mm:ss")

' ---------------------- Delivery Dates ----------------------
If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
End If

' ---------------------- Payment Means ----------------------
Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
    paymentmean.InstructionNote = "Returned items"
End If
If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
End If
If Len(Trim$(paymentnote & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
End If
inv.PaymentMeans.Add paymentmean

' ---------------------- Supplier ----------------------
inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
inv.SupplierParty.PartyIdentification.schemeID = cschemeID

inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
Dim supplierCountry As String
supplierCountry = Trim$(cIdentificationCode & "")
If supplierCountry = "" Then supplierCountry = "SA"
inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry
inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")

' ---------------------- Customer (Standard/Simplified) ----------------------
Dim isStandard As Boolean
isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")

Dim isSimplified As Boolean
isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
               And (LCase$(mTableName) = "tbleinvoice" _
                 Or LCase$(mTableName) = "propertyduebatchdetail" _
                 Or LCase$(mTableName) = "debitandcreditnotification")

If isStandard Or isSimplified Then
    Dim CUSTID As String
    If Len(Trim$(Id700 & "")) > 0 Then
        CUSTID = Trim$(Id700 & "")
    Else
        CUSTID = Trim$(Identificationid & "")
    End If
    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
        CUSTID = "300000000000003": schemeID = "OTH"
    End If

    inv.CustomerParty.PartyIdentification.ID = CUSTID
    If Len(Trim$(schemeID & "")) = 0 Then schemeID = "OTH"
    inv.CustomerParty.PartyIdentification.schemeID = schemeID

    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
    cAddStreet = Trim$(AdditionalStreetName & "")
    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
    cSubent = Trim$(CountrySubentity & "")
    cSubCity = Trim$(CitySubdivisionName & "")
    cCountry = Trim$(IdentificationCode & "")
    If Len(cCountry) = 0 Then cCountry = "SA"

    inv.CustomerParty.PostalAddress.StreetName = cStreet
    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
    inv.CustomerParty.PostalAddress.CityName = cCity
    inv.CustomerParty.PostalAddress.PostalZone = cPost
    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry

    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
    If Len(Trim$(Id700 & "")) > 0 Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
    End If
End If

' =========================================================
' ============   «·»‰Êœ +  Ã„Ì⁄ «·Œ’Ê„«    ===============
' =========================================================
Dim StrSQL As String
StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT,Trans_DiscountType,Trans_Discount, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  TD.ItemDiscountType, "
StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText

' *** „ €Ì—«  ·· Ã„Ì⁄ »œﬁ… ⁄«·Ì… ***
Dim SumExNoDisc As Double                 ' ≈Ã„«·Ì «·√”«” ﬁ»· √Ì Œ’„ Ê÷—Ì»… (Ex-VAT ﬁ»· «·Œ’„)
Dim SumTaxableExNoDisc As Double          ' ≈Ã„«·Ì «·√”«” «·Œ«÷⁄ ﬁ»· «·Œ’„
Dim SumExemptExNoDisc As Double           ' ≈Ã„«·Ì «·√”«” «·„⁄›Ï ﬁ»· «·Œ’„
Dim SumLineDiscEx As Double               ' „Ã„Ê⁄ Œ’„ «·”ÿÊ— »⁄œ  ÕÊÌ·Â Ex-VAT
Dim mTrans_DiscountType As Integer
Dim mTrans_Discount As Double

SumExNoDisc = 0#
SumTaxableExNoDisc = 0#
SumExemptExNoDisc = 0#
SumLineDiscEx = 0#

If Not rs.EOF Then
    mTrans_DiscountType = val(rs!Trans_DiscountType & "")
    mTrans_Discount = val(rs!Trans_Discount & "")
End If

Do Until rs.EOF
    Dim lineEx As Double
    Dim Rate As Double
    Dim Qty As Double
    Dim basePrice As Double

    Dim discType As Long
    Dim discPercent As Double
    Dim discValue As Double
    Dim discRaw As Double
    Dim lineDiscEx As Double
    Dim lineBaseForPercent As Double
    
    ' ﬁÌ„ «·”ÿ— «·√”«”Ì…
    lineEx = val(rs!LineEx_NoDisc & "")   ' ﬁÌ„… «·”ÿ— ﬁ»· «·Œ’„ Ex-VAT
    Rate = val(rs!vatRate & "")
    Qty = val(rs!ShowQty & "")
    basePrice = val(rs!BasePrice_Ex & "")

    ' »Ì«‰«  «·Œ’„
    discType = val(rs!ItemDiscountType & "")
    discPercent = val(rs!ItemDiscount & "")   '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ‰”»…
    discValue = val(rs!ItemDiscount & "")     '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ﬁÌ„…
    discRaw = 0

    ' ===== Õ”«» «·Œ’„ «·„Œ“‰ ⁄·Ï «·”ÿ— (»«·√”⁄«— «· Ì √œŒ·Â« «·„” Œœ„) =====
    Select Case discType
        Case 3      ' Œ’„ ‰”»… %
            ' ·Ê «·√”⁄«— »«·÷—Ì»… ? ‰”» Â „‰ ﬁÌ„… «·”ÿ— »«·÷—Ì»…
            ' ·Ê «·√”⁄«— »œÊ‰ ÷—Ì»… ? ‰”» Â „‰ Ex-VAT
            If SystemOptions.PriceWithVAT And Rate > 0 Then
                ' lineEx ÂÊ Ex-VAT° ‰ÕÊ·Â ≈·Ï ‘«„· ÷—Ì»…
                lineBaseForPercent = lineEx * (1 + Rate / 100#)
            Else
                lineBaseForPercent = lineEx
            End If

            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) _
                      + val(rs!TotalDiscountPerLine & "")

        Case 2      ' Œ’„ ﬁÌ„…
            discRaw = discValue + val(rs!TotalDiscountPerLine & "")

        Case Else   ' „›Ì‘ ‰Ê⁄ Œ’„ „Õœœ ? ‰ﬂ ›Ì »‹ TotalDiscountPerLine ·Ê „ÊÃÊœ
            discRaw = val(rs!TotalDiscountPerLine & "")
    End Select

    ' =====  ÕÊÌ· «·Œ’„ ≈·Ï Ex-VAT =====
    If SystemOptions.PriceWithVAT And Rate > 0 Then
        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
    Else
        lineDiscEx = Round(discRaw, 6)
    End If

    ' =====  Ã„Ì⁄ «·√”«”«  ﬁ»· «·Œ’„ =====
    SumExNoDisc = SumExNoDisc + lineEx

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
    Else
        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
    End If

    ' =====  Ã„Ì⁄ Œ’Ê„«  «·”ÿÊ— ﬂ‹ "Œ’„ ﬂ·Ì" ⁄·Ï «·›« Ê—… =====
    SumLineDiscEx = SumLineDiscEx + lineDiscEx

    ' ===== ≈‰‘«¡ ”ÿ— UBL »·« Œ’„ ”ÿ—Ì =====
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    If docType = 0 Then
        invline.Price.EncludingVat = SystemOptions.PriceWithVAT
    Else
        invline.Price.EncludingVat = False
    End If
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = rs!ItemName & ""

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = Rate
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = Rate
    Else
        invline.Item.classifiedTaxCategory.ID = "E"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
    End If

    ' ·« ‰÷Ì› AllowanceCharge ⁄·Ï „” ÊÏ «·”ÿ— ≈ÿ·«ﬁ«
    inv.InvoiceLines.Add invline

    rs.MoveNext
Loop

rs.Close

' =========================================================
' ==============  Œ’„ „ÊÕ¯œ (”ÿÊ— + ÂÌœ—)  ================
' =========================================================
Dim HeaderDiscEx As Double

' SumLineDiscEx Ã«Ì „‰ «··Ê» : „Ã„Ê⁄ Œ’Ê„«  «·”ÿÊ— »⁄œ  ÕÊÌ·Â« Ex-VAT

Dim TransDiscountRaw As Double

' SumLineDiscEx Ã«Ì „‰ «··Ê» : „Ã„Ê⁄ Œ’Ê„«  «·”ÿÊ— »⁄œ  ÕÊÌ·Â« Ex-VAT
HeaderDiscEx = Round(SumLineDiscEx, 2)

' ===== „⁄«·Ã… Œ’„ —√” «·›« Ê—… Trans_Discount =====
TransDiscountRaw = val(Trans_Discount & "")

If TransDiscountRaw <> 0 Then
    ' «· Õﬁﬁ: Â· Trans_Discount ‘«„· ÷—Ì»…ø
    ' Õ”» Trans_DiscountType:
    ' - ≈–« ﬂ«‰ Type = 1 (‰”»… %) -> ÌÃ»  ÿ»ÌﬁÂ ⁄·Ï «·√”«” «·„‰«”»
    ' - ≈–« ﬂ«‰ Type = 2 (ﬁÌ„…) -> ÌÃ» «· Õﬁﬁ „‰ SystemOptions.PriceWithVAT
    
    Dim TransDiscountEx As Double
    
    If val(mTrans_DiscountType) = 1 Then
        ' Œ’„ ‰”»… % - Ìÿ»ﬁ ⁄·Ï «·„Ã„Ê⁄
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            ' «·Œ’„ ⁄·Ï «·√”«” ‘«„· «·÷—Ì»…
            Dim TotalWithVAT As Double
            TotalWithVAT = SumExNoDisc * (1 + 0.15)
            TransDiscountRaw = Round(TotalWithVAT * TransDiscountRaw / 100#, 6)
            '  ÕÊÌ· ≈·Ï Ex-VAT
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            ' «·Œ’„ ⁄·Ï «·√”«” »œÊ‰ ÷—Ì»…
            TransDiscountRaw = Round(SumExNoDisc * TransDiscountRaw / 100#, 6)
            TransDiscountEx = TransDiscountRaw
        End If
    Else
        ' Œ’„ ﬁÌ„… À«» …
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            ' «·ﬁÌ„… «·„œŒ·… ‘«„·… «·÷—Ì»… - ‰ÕÊ·Â« ≈·Ï Ex-VAT
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            ' «·ﬁÌ„… «·„œŒ·… »œÊ‰ ÷—Ì»…
            TransDiscountEx = TransDiscountRaw
        End If
    End If
    
    ' ≈÷«›… «·Œ’„ «·„ÕÊ¯· ≈·Ï Ex-VAT
    HeaderDiscEx = HeaderDiscEx + TransDiscountEx
End If

If HeaderDiscEx < 0 Then HeaderDiscEx = 0

'  Ê“Ì⁄ œ«Œ·Ì (‰ˆ”»Ì) ·„⁄—›…  √ÀÌ— «·Œ’„ ⁄·Ï «·Œ«÷⁄/«·„⁄›Ï ·Õ”«» VAT
Dim BaseTotal As Double: BaseTotal = SumExNoDisc
Dim Allow_S As Double, Allow_E As Double

If BaseTotal > 0 Then
    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 2)  '  ﬁ—Ì» Â‰«
    Allow_E = Round(HeaderDiscEx - Allow_S, 2)                            '  ﬁ—Ì» Â‰«
Else
    Allow_S = 0
    Allow_E = HeaderDiscEx
End If

' ===== Õ”«» Œ’„ «·ÂÌœ— ‘«„·« «·÷—Ì»… =====
Dim HeaderDiscVat As Double         ' ÷—Ì»… «·Œ’„ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… ›ﬁÿ
Dim Allow_S_Inc As Double           ' Œ’„ Œ«÷⁄ ‘«„· ÷—Ì»…
Dim Allow_E_Inc As Double           ' Œ’„ „⁄›Ï (ÂÊ ‰›”Â Ex-VAT)

' ‰› —÷ √‰ «·÷—Ì»… «·„ÊÕœ… 15% ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S
HeaderDiscVat = Round(Allow_S * 0.15, 2)

Allow_S_Inc = Round(Allow_S, 2)
Allow_E_Inc = Round(Allow_E, 2)

' --------- √÷› AllowanceCharge ··ÂÌœ— (Ê«Õœ √Ê «À‰Ì‰) ---------
If HeaderDiscEx > 0# Then

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S (‘«„· «·÷—Ì»…)
    If Allow_S_Inc > 0# Then
        Dim allowanceS As Object
        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceS.ChargeIndicator = False
        Allow_S_Inc = Allow_S_Inc
        allowanceS.Amount = Format(Allow_S_Inc, "0.##")          ' Œ’„ Œ«÷⁄ ‘«„· VAT
        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceS.taxCategory.ID = "S"
        allowanceS.taxCategory.Percent = 15#
        inv.allowanceCharges.Add allowanceS
    End If

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·„⁄›«… E (»œÊ‰ ÷—Ì»…)
    If Allow_E_Inc > 0# Then
        Dim allowanceE As Object
        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceE.ChargeIndicator = False
        allowanceE.Amount = Format(Allow_E_Inc, "0.##")          ' Œ’„ „⁄›Ï Ex-VAT
        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceE.taxCategory.ID = "E"
        allowanceE.taxCategory.Percent = 0#
        inv.allowanceCharges.Add allowanceE
    End If

End If


' =========================================================
' ====================   «·≈Ã„«·Ì«    =====================
' =========================================================

' ***  ﬁ—Ì» «·„Ã«„Ì⁄ ·„‰“· Ì‰ ⁄‘—Ì Ì‰ „»ﬂ—« ***
SumExNoDisc = Round(SumExNoDisc, 2)
SumTaxableExNoDisc = Round(SumTaxableExNoDisc, 2)
HeaderDiscEx = Round(HeaderDiscEx, 2)

' LineExtensionAmount = √”«” «·»‰Êœ ﬁ»· √Ì Œ’„ Ê÷—Ì»… (Ex-VAT ﬁ»· «·Œ’„)
inv.LegalMonetaryTotal.LineExtensionAmount = _
    Format(SumExNoDisc, "0.##")

' AllowanceTotalAmount = ≈Ã„«·Ì «·Œ’„ «·„ÊÕ¯œ (”ÿÊ— + ÂÌœ—) Ex-VAT
inv.LegalMonetaryTotal.AllowanceTotalAmount = _
    Format(HeaderDiscEx + HeaderDiscVat, "0.##")

' TaxExclusiveAmount = LineExtension - ≈Ã„«·Ì «·Œ’„ Ex-VAT (Œ«÷⁄ + „⁄›Ï)
Dim TaxExclusiveAfterDisc As Double
TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 2)  '  ﬁ—Ì» Â‰«
If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0

inv.LegalMonetaryTotal.TaxExclusiveAmount = _
    Format(TaxExclusiveAfterDisc, "0.##")

' ===== «Õ ”«» ÷—Ì»… «·ﬁÌ„… «·„÷«›… „⁄ „—«⁄«… «·Œ’„ =====
Dim TaxableAfterDisc As Double
Dim TotalVat As Double

' «·√”«” «·Œ«÷⁄ »⁄œ Œ’„ «·Ã“¡ «·Œ«÷⁄ „‰ «·Œ’„ «·ﬂ·Ì (Allow_S)
TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 2)  ' ***  ﬁ—Ì» Â‰« ·‹ 2 „‰«“· ***
If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0

' VAT = 15% „‰ «·√”«” «·Œ«÷⁄ »⁄œ «·Œ’„
'TotalVat = Round(TaxableAfterDisc * 0.15, 2)
Dim rawVat As Double
rawVat = TaxableAfterDisc * 0.15 ' ”Ì⁄ÿÌ 4716.075
' ‰” Œœ„ Fix ·≈“«·… „« »⁄œ «·⁄·«„… «·⁄‘—Ì… »⁄œ «·÷—» ›Ì 100
TotalVat = Fix(rawVat * 100) / 100 ' (Fix(471607.5) / 100) = 471607 / 100 = 4716.07
' --- ‰Â«Ì… «· ⁄œÌ· ---


inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")

' *** TaxInclusive = TaxExclusive + VAT »œﬁ… ⁄«·Ì… ·„‰⁄ «· —«ﬂ„ ***
Dim TaxInclusive As Double
' «” Œœ«„ CDbl + Format ·÷„«‰ œﬁ… «· ﬁ—Ì»
TaxInclusive = CDbl(Format(TaxExclusiveAfterDisc, "0.00")) + CDbl(Format(TotalVat, "0.00"))
TaxInclusive = Round(TaxInclusive, 2)

inv.LegalMonetaryTotal.TaxInclusiveAmount = _
    Format(TaxInclusive, "0.##")

' «·ÕﬁÊ· «·√Œ—Ï
inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"

' PrepaidAmount (·Ê ›« Ê—… ”œ«œ „ﬁœ¯„ 388 ›ﬁÿ)
Dim Prepaid As Double
If val(InvoiceTypeCodeID) = 388 Then
    Prepaid = val(PrepaidAmount & "")
Else
    Prepaid = 0
End If
inv.LegalMonetaryTotal.PrepaidAmount = _
    Format(Round(Prepaid, 2), "0.##")

' Payable = TaxInclusive - Prepaid („⁄ ÷»ÿ «· ﬁ—Ì»)
Dim PayableAmountCalc As Double
PayableAmountCalc = Round(TaxInclusive - Prepaid, 2)
If PayableAmountCalc < 0 Then PayableAmountCalc = 0
inv.LegalMonetaryTotal.PayableAmount = _
    Format(PayableAmountCalc, "0.##")

' =========================================================
' ==================   «· ÊﬁÌ⁄ Ê«·≈—”«·   =================
' =========================================================
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey

If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)

If res.IsValid = True Then
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    Dim apireqlogic As Object, tokenresponse As Object
    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")

    invrequestbody.invoice = res.EncodedInvoice
    invrequestbody.InvoiceHash = res.InvoiceHash
    invrequestbody.UUID = res.UUID

    Select Case SendingMode
        Case 0 ' DEV (Compliance)
            apireqlogic.SetApiRequestLogic 0
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                    updatedatatable val(docType), mTableName, mFieldIDName
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 1 ' SIMU (Reporting/Clearance)
            apireqlogic.SetApiRequestLogic 1
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 2 ' PROD (Clearance)
            apireqlogic.SetApiRequestLogic 2
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If
    End Select

    GenerateInvoice_ZatcaV3DD = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV3DD = False
End If

End Function


'
Public Function GenerateInvoice_ZatcaV4( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "Transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0 _
) As Boolean
On Error Resume Next

' =====================  Õ÷Ì— «·‹ UBL/Invoice =====================
Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

' ---- « ’«· (ﬂ„« ÂÊ »œÊ‰  ⁄œÌ·) ----
If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If

' ---- »Ì«‰«  «·‘—ﬂ…/«·›—⁄ ----
If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

' ---------------------- BillingReference (CN/DN only) ----------------------
If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If
End If

' ---------------------- ‰Ê⁄ «·›« Ê—… ----------------------
inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
inv.InvoiceTypeCode.Name = InvoiceTypeCodename

' ---------------------- ICV / UUID / PIH ----------------------
inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)

Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
icvOnlyDigits = ""
icvStr = Trim$(uuidCounter & "")
If icvStr = "" Then icvStr = "1"
For i = 1 To Len(icvStr)
    Ch = mId$(icvStr, i, 1)
    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
Next i
If icvOnlyDigits = "" Then icvOnlyDigits = "1"
inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits

inv.ID = ID

' ---------------------- Issue Date & Time ----------------------
If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")

If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
' „Â„: mm = œﬁ«∆ﬁ / ss = ÀÊ«‰Ì
inv.IssueTime = Format(IssueTim, "HH:mm:ss")

' ---------------------- Delivery Dates ----------------------
If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
End If

' ---------------------- Payment Means ----------------------
Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
    paymentmean.InstructionNote = "Returned items"
End If
If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
End If
If Len(Trim$(paymentnote & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
End If
inv.PaymentMeans.Add paymentmean

' ---------------------- Supplier ----------------------
inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
inv.SupplierParty.PartyIdentification.schemeID = cschemeID

inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
Dim supplierCountry As String
supplierCountry = Trim$(cIdentificationCode & "")
If supplierCountry = "" Then supplierCountry = "SA"
inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry
inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")

' ---------------------- Customer (Standard/Simplified) ----------------------
Dim isStandard As Boolean
isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")

Dim isSimplified As Boolean
isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
               And (LCase$(mTableName) = "tblEinvoice" _
                 Or LCase$(mTableName) = "propertyduebatchdetail" _
                 Or LCase$(mTableName) = "debitandcreditnotification")

If isStandard Or isSimplified Then
    Dim CUSTID As String
    If Len(Trim$(Id700 & "")) > 0 Then
        CUSTID = Trim$(Id700 & "")
    Else
        CUSTID = Trim$(Identificationid & "")
    End If
    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
        CUSTID = "300000000000003": schemeID = "OTH"
    End If

    inv.CustomerParty.PartyIdentification.ID = CUSTID
    If Len(Trim$(schemeID & "")) = 0 Then schemeID = "OTH"
    inv.CustomerParty.PartyIdentification.schemeID = schemeID

    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
    cAddStreet = Trim$(AdditionalStreetName & "")
    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
    cSubent = Trim$(CountrySubentity & "")
    cSubCity = Trim$(CitySubdivisionName & "")
    cCountry = Trim$(IdentificationCode & "")
    If Len(cCountry) = 0 Then cCountry = "SA"

    inv.CustomerParty.PostalAddress.StreetName = cStreet
    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
    inv.CustomerParty.PostalAddress.CityName = cCity
    inv.CustomerParty.PostalAddress.PostalZone = cPost
    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry

    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
    If Len(Trim$(Id700 & "")) > 0 Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
    End If
End If

' =========================================================
' ============   «·»‰Êœ +  Ã„Ì⁄ «·Œ’Ê„«    ===============
' =========================================================
Dim StrSQL As String
StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & " t.Trans_Discount,t.Trans_DiscountType, I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  LineDiscRaw = ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0), "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID




StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT,Trans_DiscountType,Trans_Discount, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  TD.ItemDiscountType, "
StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
Dim SumExNoDisc As Double                 ' ≈Ã„«·Ì «·√”«” ﬁ»· √Ì Œ’„ (Ex-VAT)
Dim SumTaxableExNoDisc As Double          ' ≈Ã„«·Ì «·√”«” «·Œ«÷⁄ ﬁ»· «·Œ’„
Dim SumExemptExNoDisc As Double           ' ≈Ã„«·Ì «·√”«” «·„⁄›Ï ﬁ»· «·Œ’„
Dim SumLineDiscEx As Double               ' „Ã„Ê⁄ Œ’„ «·”ÿÊ— »⁄œ  ÕÊÌ·Â Ex-VAT
Dim mTrans_DiscountType As Integer
Dim mTrans_Discount As Double
If Not rs.EOF Then
    mTrans_DiscountType = val(rs!Trans_DiscountType & "")
    mTrans_Discount = val(rs!Trans_Discount & "")
End If
Do Until rs.EOF
    Dim lineEx As Double
    Dim Rate As Double
    Dim Qty As Double
    Dim basePrice As Double

    Dim discType As Long
    Dim discPercent As Double
    Dim discValue As Double
    Dim discRaw As Double
    Dim lineDiscEx As Double
    Dim lineBaseForPercent As Double
    
    ' ﬁÌ„ «·”ÿ— «·√”«”Ì…
    lineEx = val(rs!LineEx_NoDisc & "")   ' ﬁÌ„… «·”ÿ— ﬁ»· «·Œ’„ Ex-VAT
    Rate = val(rs!vatRate & "")
    Qty = val(rs!ShowQty & "")
    basePrice = val(rs!BasePrice_Ex & "")

    ' »Ì«‰«  «·Œ’„
    discType = val(rs!ItemDiscountType & "")
    discPercent = val(rs!ItemDiscount & "")   '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ‰”»…
    discValue = val(rs!ItemDiscount & "")     '  ı” Œœ„ ·Ê ‰Ê⁄ «·Œ’„ ﬁÌ„…
    discRaw = 0

    ' ===== Õ”«» «·Œ’„ «·„Œ“‰ ⁄·Ï «·”ÿ— (»«·√”⁄«— «· Ì √œŒ·Â« «·„” Œœ„) =====
    Select Case discType
        Case 3      ' Œ’„ ‰”»… %
            ' ·Ê «·√”⁄«— »«·÷—Ì»… ? ‰”» Â „‰ ﬁÌ„… «·”ÿ— »«·÷—Ì»…
            ' ·Ê «·√”⁄«— »œÊ‰ ÷—Ì»… ? ‰”» Â „‰ Ex-VAT
            If SystemOptions.PriceWithVAT And Rate > 0 Then
                ' lineEx ÂÊ Ex-VAT° ‰ÕÊ·Â ≈·Ï ‘«„· ÷—Ì»…
                lineBaseForPercent = lineEx * (1 + Rate / 100#)
            Else
                lineBaseForPercent = lineEx
            End If

            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) _
                      + val(rs!TotalDiscountPerLine & "")

        Case 2      ' Œ’„ ﬁÌ„…
            discRaw = discValue + val(rs!TotalDiscountPerLine & "")

        Case Else   ' „›Ì‘ ‰Ê⁄ Œ’„ „Õœœ ? ‰ﬂ ›Ì »‹ TotalDiscountPerLine ·Ê „ÊÃÊœ
            discRaw = val(rs!TotalDiscountPerLine & "")
    End Select

    ' =====  ÕÊÌ· «·Œ’„ ≈·Ï Ex-VAT =====
    If SystemOptions.PriceWithVAT And Rate > 0 Then
        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
    Else
        lineDiscEx = Round(discRaw, 6)
    End If

    ' =====  Ã„Ì⁄ «·√”«”«  ﬁ»· «·Œ’„ =====
    SumExNoDisc = SumExNoDisc + lineEx

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
    Else
        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
    End If

    ' =====  Ã„Ì⁄ Œ’Ê„«  «·”ÿÊ— ﬂ‹ "Œ’„ ﬂ·Ì" ⁄·Ï «·›« Ê—… =====
    SumLineDiscEx = SumLineDiscEx + lineDiscEx

    ' ===== ≈‰‘«¡ ”ÿ— UBL »·« Œ’„ ”ÿ—Ì =====
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    invline.Price.EncludingVat = False
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = rs!ItemName & ""

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = Rate
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = Rate
    Else
        invline.Item.classifiedTaxCategory.ID = "E"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
    End If

    ' ·« ‰÷Ì› AllowanceCharge ⁄·Ï „” ÊÏ «·”ÿ— ≈ÿ·«ﬁ«
    inv.InvoiceLines.Add invline

    rs.MoveNext
Loop

rs.Close

' =========================================================
' ==============  Œ’„ „ÊÕ¯œ (”ÿÊ— + ÂÌœ—)  ================
' =========================================================
Dim HeaderDiscEx As Double

' SumLineDiscEx Ã«Ì „‰ «··Ê» : „Ã„Ê⁄ Œ’Ê„«  «·”ÿÊ— »⁄œ  ÕÊÌ·Â« Ex-VAT

Dim TransDiscountRaw As Double

' SumLineDiscEx Ã«Ì „‰ «··Ê» : „Ã„Ê⁄ Œ’Ê„«  «·”ÿÊ— »⁄œ  ÕÊÌ·Â« Ex-VAT
HeaderDiscEx = Round(SumLineDiscEx, 2)

' ===== „⁄«·Ã… Œ’„ —√” «·›« Ê—… Trans_Discount =====
TransDiscountRaw = val(Trans_Discount & "")

If TransDiscountRaw <> 0 Then
    ' «· Õﬁﬁ: Â· Trans_Discount ‘«„· ÷—Ì»…ø
    ' Õ”» Trans_DiscountType:
    ' - ≈–« ﬂ«‰ Type = 1 (‰”»… %) -> ÌÃ»  ÿ»ÌﬁÂ ⁄·Ï «·√”«” «·„‰«”»
    ' - ≈–« ﬂ«‰ Type = 2 (ﬁÌ„…) -> ÌÃ» «· Õﬁﬁ „‰ SystemOptions.PriceWithVAT
    
    Dim TransDiscountEx As Double
    
    If val(mTrans_DiscountType) = 1 Then
        ' Œ’„ ‰”»… % - Ìÿ»ﬁ ⁄·Ï «·„Ã„Ê⁄
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            ' «·Œ’„ ⁄·Ï «·√”«” ‘«„· «·÷—Ì»…
            Dim TotalWithVAT As Double
            TotalWithVAT = SumExNoDisc * (1 + 0.15)
            TransDiscountRaw = Round(TotalWithVAT * TransDiscountRaw / 100#, 6)
            '  ÕÊÌ· ≈·Ï Ex-VAT
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            ' «·Œ’„ ⁄·Ï «·√”«” »œÊ‰ ÷—Ì»…
            TransDiscountRaw = Round(SumExNoDisc * TransDiscountRaw / 100#, 6)
            TransDiscountEx = TransDiscountRaw
        End If
    Else
        ' Œ’„ ﬁÌ„… À«» …
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            ' «·ﬁÌ„… «·„œŒ·… ‘«„·… «·÷—Ì»… - ‰ÕÊ·Â« ≈·Ï Ex-VAT
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            ' «·ﬁÌ„… «·„œŒ·… »œÊ‰ ÷—Ì»…
            TransDiscountEx = TransDiscountRaw
        End If
    End If
    
    ' ≈÷«›… «·Œ’„ «·„ÕÊ¯· ≈·Ï Ex-VAT
    HeaderDiscEx = HeaderDiscEx + TransDiscountEx
End If

If HeaderDiscEx < 0 Then HeaderDiscEx = 0

'  Ê“Ì⁄ œ«Œ·Ì (‰ˆ”»Ì) ·„⁄—›…  √ÀÌ— «·Œ’„ ⁄·Ï «·Œ«÷⁄/«·„⁄›Ï ·Õ”«» VAT
Dim BaseTotal As Double: BaseTotal = SumExNoDisc
Dim Allow_S As Double, Allow_E As Double

If BaseTotal > 0 Then
    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 6)
    Allow_E = Round(HeaderDiscEx - Allow_S, 6)
Else
    Allow_S = 0
    Allow_E = HeaderDiscEx
End If

' ===== Õ”«» Œ’„ «·ÂÌœ— ‘«„·« «·÷—Ì»… =====
Dim HeaderDiscVat As Double         ' ÷—Ì»… «·Œ’„ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… ›ﬁÿ
Dim Allow_S_Inc As Double           ' Œ’„ Œ«÷⁄ ‘«„· ÷—Ì»…
Dim Allow_E_Inc As Double           ' Œ’„ „⁄›Ï (ÂÊ ‰›”Â Ex-VAT)

' ‰› —÷ √‰ «·÷—Ì»… «·„ÊÕœ… 15% ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S
HeaderDiscVat = Round(Allow_S * 0.15, 6)

'Allow_S_Inc = Round(Allow_S + HeaderDiscVat, 2)
Allow_S_Inc = Round(Allow_S, 2)
Allow_E_Inc = Round(Allow_E, 2)

' --------- √÷› AllowanceCharge ··ÂÌœ— (Ê«Õœ √Ê «À‰Ì‰) ---------
If HeaderDiscEx > 0# Then

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·Œ«÷⁄… S (‘«„· «·÷—Ì»…)
    If Allow_S_Inc > 0# Then
        Dim allowanceS As Object
        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceS.ChargeIndicator = False
        Allow_S_Inc = Allow_S_Inc
        allowanceS.Amount = Format(Allow_S_Inc, "0.##")          ' Œ’„ Œ«÷⁄ ‘«„· VAT
        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceS.taxCategory.ID = "S"
        allowanceS.taxCategory.Percent = 15#
        inv.allowanceCharges.Add allowanceS
    End If

    ' Ã“¡ «·Œ’„ «·Ê«ﬁ⁄ ⁄·Ï «·»‰Êœ «·„⁄›«… E (»œÊ‰ ÷—Ì»…)
    If Allow_E_Inc > 0# Then
        Dim allowanceE As Object
        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceE.ChargeIndicator = False
        allowanceE.Amount = Format(Allow_E_Inc, "0.##")          ' Œ’„ „⁄›Ï Ex-VAT
        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceE.taxCategory.ID = "E"
        allowanceE.taxCategory.Percent = 0#
        inv.allowanceCharges.Add allowanceE
    End If

End If


' =========================================================
' ====================   «·≈Ã„«·Ì«    =====================
' =========================================================

' LineExtensionAmount = √”«” «·»‰Êœ ﬁ»· √Ì Œ’„ Ê÷—Ì»… (Ex-VAT ﬁ»· «·Œ’„)
inv.LegalMonetaryTotal.LineExtensionAmount = _
    Format(Round(SumExNoDisc, 2), "0.##")

' AllowanceTotalAmount = ≈Ã„«·Ì «·Œ’„ «·„ÊÕ¯œ (”ÿÊ— + ÂÌœ—) Ex-VAT
inv.LegalMonetaryTotal.AllowanceTotalAmount = _
    Format(Round(HeaderDiscEx, 2), "0.##")

' TaxExclusiveAmount = LineExtension - ≈Ã„«·Ì «·Œ’„ Ex-VAT (Œ«÷⁄ + „⁄›Ï)
Dim TaxExclusiveAfterDisc As Double
TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 6)
If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0

inv.LegalMonetaryTotal.TaxExclusiveAmount = _
    Format(Round(TaxExclusiveAfterDisc, 2), "0.##")

' ===== «Õ ”«» ÷—Ì»… «·ﬁÌ„… «·„÷«›… „⁄ „—«⁄«… «·Œ’„ =====
Dim TaxableAfterDisc As Double
Dim TotalVat As Double

' «·√”«” «·Œ«÷⁄ »⁄œ Œ’„ «·Ã“¡ «·Œ«÷⁄ „‰ «·Œ’„ «·ﬂ·Ì (Allow_S)
TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 6)
If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0

' VAT = 15% „‰ «·√”«” «·Œ«÷⁄ »⁄œ «·Œ’„
TotalVat = Round(TaxableAfterDisc * 0.15, 2)

inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")

' TaxInclusive = TaxExclusive + VAT
Dim TaxInclusive As Double
TaxInclusive = Round(TaxExclusiveAfterDisc + TotalVat, 2)
inv.LegalMonetaryTotal.TaxInclusiveAmount = _
    Format(TaxInclusive, "0.##")

' «·ÕﬁÊ· «·√Œ—Ï
inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"

' PrepaidAmount (·Ê ›« Ê—… ”œ«œ „ﬁœ¯„ 388 ›ﬁÿ)
Dim Prepaid As Double
If val(InvoiceTypeCodeID) = 388 Then
    Prepaid = val(PrepaidAmount & "")
Else
    Prepaid = 0
End If
inv.LegalMonetaryTotal.PrepaidAmount = _
    Format(Round(Prepaid, 2), "0.##")

' Payable = TaxInclusive - Prepaid („⁄ ÷»ÿ «· ﬁ—Ì»)
Dim PayableAmountCalc As Double
PayableAmountCalc = Round(TaxInclusive - Prepaid, 2)
If PayableAmountCalc < 0 Then PayableAmountCalc = 0
inv.LegalMonetaryTotal.PayableAmount = _
    Format(PayableAmountCalc, "0.##")

' =========================================================
' ==================   «· ÊﬁÌ⁄ Ê«·≈—”«·   =================
' =========================================================
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey

If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)

If res.IsValid = True Then
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    Dim apireqlogic As Object, tokenresponse As Object
    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")

    invrequestbody.invoice = res.EncodedInvoice
    invrequestbody.InvoiceHash = res.InvoiceHash
    invrequestbody.UUID = res.UUID

    Select Case SendingMode
        Case 0 ' DEV (Compliance)
            apireqlogic.SetApiRequestLogic 0
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                    updatedatatable val(docType), mTableName, mFieldIDName
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 1 ' SIMU (Reporting/Clearance)
            apireqlogic.SetApiRequestLogic 1
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 2 ' PROD (Clearance)
            apireqlogic.SetApiRequestLogic 2
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If
    End Select

    GenerateInvoice_ZatcaV4 = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV4 = False
End If

End Function


Public Function GenerateInvoice_ZatcaV3Work( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "Transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0 _
) As Boolean
On Error Resume Next

' =====================  Õ÷Ì— «·‹ UBL/Invoice =====================
Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")


If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If

' »Ì«‰«  «·‘—ﬂ…/«·›—⁄ Ê«·⁄„Ì· (‰›” —Ê Ì‰ﬂ)
If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

' Header basics
'---------------------- BillingReference (CN/DN only) ----------------------
If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If
End If
    inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
  inv.InvoiceTypeCode.Name = InvoiceTypeCodename
'---------------------- ICV / UUID ----------------------
' ZATCA Ì‘ —ÿ ICV √—ﬁ«„ ›ﬁÿ
 inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
Dim icvStr As String
icvOnlyDigits = ""
icvStr = Trim$(uuidCounter & "")
If icvStr = "" Then icvStr = "1"
' ≈‰ ·„ Ìﬂ‰ √—ﬁ«„ ›ﬁÿ° ‰ÕÊ¯·Â ·√—ﬁ«„ »≈“«·… √Ì „Õ«—› €Ì— —ﬁ„Ì…
Dim i As Long, Ch As String
For i = 1 To Len(icvStr)
    Ch = mId$(icvStr, i, 1)
    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
Next i
If icvOnlyDigits = "" Then icvOnlyDigits = "1"
inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits
  inv.ID = ID
 '---------------------- Issue Date & Time ----------------------
If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then
    IssueDate = Date
End If
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")  ' „À«·: 2021-01-05

If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then
    IssueTim = Now
End If
inv.IssueTime = Format(IssueTim, "HH:mm:ss")     ' „À«·: 09:32:40

'---------------------- Delivery Dates ----------------------
If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
End If
    inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
  inv.InvoiceTypeCode.Name = InvoiceTypeCodename
'---------------------- Payment Means ----------------------
Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")

If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
    paymentmean.InstructionNote = "Returned items"
End If

If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
End If
If Len(Trim$(paymentnote & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
End If

inv.PaymentMeans.Add paymentmean

'---------------------- Supplier (AccountingSupplierParty) ----------------------
inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
inv.SupplierParty.PartyIdentification.schemeID = cschemeID

inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")
Dim supplierCountry As String
supplierCountry = Trim$(cIdentificationCode & "")
If supplierCountry = "" Then supplierCountry = "SA"
inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry

inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")

'---------------------- Customer (AccountingCustomerParty) ----------------------
Dim isStandard As Boolean
isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")

Dim isSimplified As Boolean
isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
               And (LCase$(mTableName) = "tblEinvoice" _
                 Or LCase$(mTableName) = "propertyduebatchdetail" _
                 Or LCase$(mTableName) = "debitandcreditnotification")

If isStandard Or isSimplified Then
    ' PartyIdentification
    Dim CUSTID As String
    If Len(Trim$(Id700 & "")) > 0 Then
        CUSTID = Trim$(Id700 & "")
    Else
        CUSTID = Trim$(Identificationid & "")
    End If

    If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
        ' Default to non-resident/other
        CUSTID = "300000000000003"
        schemeID = "OTH"
    End If

    inv.CustomerParty.PartyIdentification.ID = CUSTID
    If Len(Trim$(schemeID & "")) = 0 Then schemeID = "OTH"
    inv.CustomerParty.PartyIdentification.schemeID = schemeID

    ' «·⁄‰Ê«‰ ñ ‰⁄Ê¯÷ «·ﬁÌ„ «·›«—€… »‹ N/A/00000 Õ”» «·Õ«Ã…
    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
    cAddStreet = Trim$(AdditionalStreetName & "")
    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
    cSubent = Trim$(CountrySubentity & "")
    cSubCity = Trim$(CitySubdivisionName & "")
    cCountry = Trim$(IdentificationCode & "")
    If Len(cCountry) = 0 Then cCountry = "SA"

    inv.CustomerParty.PostalAddress.StreetName = cStreet
    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
    inv.CustomerParty.PostalAddress.CityName = cCity
    inv.CustomerParty.PostalAddress.PostalZone = cPost
    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry

    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
    ' VAT Reg No
    If Len(Trim$(Id700 & "")) > 0 Then
        ' ≈–« «” Œœ„‰« ÂÊÌ… »œÌ·… Id700 ‰Œ·Ì CompanyID ›«—€/€Ì— „ÿ»ﬁ
        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
    End If
End If

'---------------------- Header AllowanceCharge (Œ’„ ⁄·Ï „” ÊÏ «·›« Ê—…) ----------------------
' „·«ÕŸ…: ZATCA Ì›÷¯· Œ’„ «·»‰Êœ° ·ﬂ‰ ‰ —ﬂ «·ÂÌœ— ≈–« ﬂ«‰ „ÿ·Ê»« ›Ì ”Ì‰«—ÌÊÂ« ﬂ.
'If val(allowancechargeAmount) <> 0 Then
'    Dim allowancecharge As Object
'    Set allowancecharge = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'    allowancecharge.ChargeIndicator = False
'    allowancecharge.Amount = Format(allowancechargeAmount, "0.##")
'    allowancecharge.AllowanceChargeReason = "Discount"
'
'    If chkTaxExempt = 1 Or val(TaxCategoryPercent) = 0 Then
'        allowancecharge.taxCategory.ID = "E"
'        allowancecharge.taxCategory.Percent = 0#
'    Else
'        allowancecharge.taxCategory.ID = "S"
'        allowancecharge.taxCategory.Percent = 15#
'    End If
'
'    inv.allowanceCharges.Add allowancecharge
'End If

'---------------------- LegalMonetaryTotal (Payable/Prepaid) ----------------------


' ===================== «Œ Ì«— «·« ’«· =====================
'Dim cn2 As ADODB.Connection

' ===================== SQL «·»‰Êœ (Transactions ›ﬁÿ) =====================
Dim StrSQL As String

' ===================== SQL «·»‰Êœ =====================

StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  LineDiscRaw = ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0), "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    ' «·”⁄— «·„Œ“‰ ‘«„· «·÷—Ì»… -> ‰ÕÊ· ··√”«” Ex-VAT
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    ' «·”⁄— «·„Œ“‰ €Ì— ‘«„· «·÷—Ì»…
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText

' =====================  Ã„Ì⁄ «·ﬁÌ„ =====================
Dim SumExNoDisc As Double              ' ≈Ã„«·Ì «·√”«” ﬁ»· √Ì Œ’„ (Ex-VAT)
Dim SumTaxableExNoDisc As Double       ' «·Ã“¡ «·Œ«÷⁄ ··÷—Ì»… „‰ «·√”«”
Dim SumLineDiscEx As Double            ' „Ã„Ê⁄ «·Œ’Ê„«  «·”ÿ—Ì… »⁄œ  ÕÊÌ·Â« Ex-VAT
Dim lineEx As Double, lineDiscRaw As Double, Rate As Double, Qty As Double, basePrice As Double
Dim lineDiscEx As Double, isTaxable As Boolean

Do Until rs.EOF
    lineEx = val(rs!LineEx_NoDisc & "")
    Rate = val(rs!vatRate & "")
    Qty = val(rs!ShowQty & "")
    basePrice = val(rs!BasePrice_Ex & "")
    lineDiscRaw = val(rs!lineDiscRaw & "")

    ' ÕÊ¯· Œ’„ «·”ÿ— ≈·Ï Ex-VAT ·Ê «·”⁄— √’·« ‘«„· «·÷—Ì»…
    If SystemOptions.PriceWithVAT And Rate > 0 Then
        lineDiscEx = Round(lineDiscRaw / (1 + (Rate / 100#)), 6)
    Else
        lineDiscEx = Round(lineDiscRaw, 6)
    End If

    '  Ã„Ì⁄
    SumExNoDisc = SumExNoDisc + lineEx
    SumLineDiscEx = SumLineDiscEx + lineDiscEx
    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
    End If

    ' ‰»‰Ì «·»‰Êœ ›Ì UBL »œÊ‰ √Ì Œ’„ ”ÿ—Ì
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    invline.Price.EncludingVat = False
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = rs!ItemName & ""

    If Rate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = Rate
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = Rate
    Else
        invline.Item.classifiedTaxCategory.ID = "E"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
    End If

    ' „Â„: ·«  ÷› √Ì AllowanceCharge ⁄·Ï „” ÊÏ «·”ÿ—
    inv.InvoiceLines.Add invline

    rs.MoveNext
Loop
rs.Close

' ===================== Œ’„ «·—√” = „Ã„Ê⁄ «·Œ’Ê„«  «·”ÿ—Ì… =====================
Dim HeaderDiscEx As Double
HeaderDiscEx = Round(SumLineDiscEx, 6)

' ·Ê ⁄‰œﬂ Œ’„ —√” ›« Ê—… ”«»ﬁ (Trans_Discount) Ê⁄«Ì“  ÷„Â:
If val(Trans_Discount & "") <> 0 Then
    ' ‰› —÷ √‰Â ﬁÌ„… (Ex-VAT). ·Ê ÂÊ ‘«„· ÷—Ì»…° ÕÊ¯·Â Ex-VAT ﬁ»· «·Ã„⁄.
    HeaderDiscEx = HeaderDiscEx + Round(val(Trans_Discount & ""), 6)
End If

' Ê“¯⁄ Œ’„ «·—√” ÷„‰Ì« ⁄·Ï «·Œ«÷⁄/«·„⁄›Ï (‰ˆ”»Ì«) ·Õ”«» VAT ’Õ
Dim TaxableShare As Double, HeaderDiscTaxable As Double
If SumExNoDisc > 0 Then
    TaxableShare = SumTaxableExNoDisc / SumExNoDisc
Else
    TaxableShare = 0
End If
HeaderDiscTaxable = Round(HeaderDiscEx * TaxableShare, 6)

' «·√”«” »⁄œ «·Œ’„
Dim TaxExclusiveAfterDisc As Double
TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 6)
If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0

' «·√”«” «·Œ«÷⁄ ··÷—Ì»… »⁄œ «·Œ’„ + «·÷—Ì»…
Dim TaxableBaseAfterDisc As Double
TaxableBaseAfterDisc = Round(SumTaxableExNoDisc - HeaderDiscTaxable, 6)
If TaxableBaseAfterDisc < 0 Then TaxableBaseAfterDisc = 0

' ‰› —÷ ÷—Ì»… 15% ﬁÌ«”Ì…
Dim TotalVat As Double
TotalVat = Round(TaxableBaseAfterDisc * 0.15, 2)

Dim TaxInclusive As Double
TaxInclusive = Round(TaxExclusiveAfterDisc + TotalVat, 2)

' ===================== AllowanceCharge ⁄·Ï —√” «·›« Ê—… =====================

'inv.LegalMonetaryTotal.PayableAmount = PayableAmount


' ===================== «·≈Ã„«·Ì«  =====================



If HeaderDiscEx > 0# Then

    Set allowancecharge = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
    allowancecharge.ChargeIndicator = False
    allowancecharge.Amount = Format(HeaderDiscEx, "0.##")
    allowancecharge.AllowanceChargeReason = "Invoice discount (summed from line discounts)"
    ' ·Ê ⁄‰œﬂ ›Ê« Ì— „⁄›«… »«·ﬂ«„·: ⁄—¯› «·›∆… "E"
    If TaxableBaseAfterDisc > 0 Then
        allowancecharge.taxCategory.ID = "S"
        allowancecharge.taxCategory.Percent = 15#
    Else
        allowancecharge.taxCategory.ID = "E"
        allowancecharge.taxCategory.Percent = 0#
    End If
    inv.allowanceCharges.Add allowancecharge
    inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(HeaderDiscEx, "0.##")        ' «·Œ’„ «·„Ã„⁄
End If


inv.LegalMonetaryTotal.LineExtensionAmount = Format(SumExNoDisc, "0.##")          ' ﬁ»· «·Œ’„

inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(TaxExclusiveAfterDisc, "0.##")
inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")
inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(TaxInclusive, "0.##")
inv.LegalMonetaryTotal.PrepaidAmount = "0.00"
If val(InvoiceTypeCodeID) = 388 Then
    inv.LegalMonetaryTotal.PrepaidAmount = PrepaidAmount
End If
inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"
inv.LegalMonetaryTotal.PayableAmount = Format(TaxInclusive, "0.##")

' ===================== «· ÊﬁÌ⁄ Ê«·≈—”«· =====================
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey
'inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(HeaderDiscEx, "0.##")        ' «·Œ’„ «·„Ã„⁄
If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)

If res.IsValid = True Then
    ' Œ“¯‰ «·‰ «∆Ã
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    ' ≈—”«· Õ”» «·‰„ÿ („·Œ¯’ Ê„ÿ«»ﬁ ·»‰Ì ﬂ)
    Dim apireqlogic As Object, tokenresponse As Object
    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")

    invrequestbody.invoice = res.EncodedInvoice
    invrequestbody.InvoiceHash = res.InvoiceHash
    invrequestbody.UUID = res.UUID

    Select Case SendingMode
        Case 0 ' DEV (Compliance)
            apireqlogic.SetApiRequestLogic 0
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                    updatedatatable val(docType), mTableName, mFieldIDName
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 1 ' SIMU (Reporting/Clearance)
            apireqlogic.SetApiRequestLogic 1
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 2 ' PROD (Clearance)
            apireqlogic.SetApiRequestLogic 2
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If
    End Select

    GenerateInvoice_ZatcaV3Work = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV3Work = False
End If

End Function

Public Function GenerateInvoice_ZatcaV35(Optional ByVal docType As Integer = 0, _
                                        Optional ByVal mTableName As String = "Transactions", _
                                        Optional ByVal mFieldIDName As String = "Transaction_ID", _
                                        Optional ByVal ActivityTypeIdInv As Integer = 0, _
                                        Optional ByVal branch_idInv As Integer = 0) As Boolean
On Error Resume Next

Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

 On Error Resume Next: Err.Clear

    ' =======================
    '  Õ÷Ì— ﬂ«∆‰«  ZATCA SDK
    ' =======================

 
    'Dim cn2 As ADODB.Connection
    If SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
        Set cn2 = POSConnection
    Else
        Set cn2 = Cn
    End If
' ------- ?????? ?????? ??????? ???? Header ?? ?? ???? -------
If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

inv.ID = ID
If IssueDate = "" Then IssueDate = Date
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")
If IssueTim = "" Then IssueTim = Now
inv.IssueTime = Format(IssueTim, "HH:MM:SS")

inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
inv.InvoiceTypeCode.Name = InvoiceTypeCodename

If DocumentCurrencyCode = "" Or DocumentCurrencyCode = "SR" Then DocumentCurrencyCode = "SAR"
If TaxCurrencyCode = "" Or TaxCurrencyCode = "SR" Then TaxCurrencyCode = "SAR"
inv.DocumentCurrencyCode = DocumentCurrencyCode
inv.TaxCurrencyCode = TaxCurrencyCode

inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
inv.AdditionalDocumentReferenceICV.UUID = uuidCounter

' ------- ??? --------
Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = PaymentMeansCode
paymentmean.PayeeFinancialAccount.ID = PayeeFinancialAccount
paymentmean.PayeeFinancialAccount.paymentnote = paymentnote
inv.PaymentMeans.Add paymentmean

' ------- SQL ?????? (Transactions) -------
Dim StrSQL As String, Percetage As Double
Dim PriceWithVAT As Boolean: PriceWithVAT = SystemOptions.PriceWithVAT

StrSQL = ""
StrSQL = StrSQL & "SELECT I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT,"
StrSQL = StrSQL & "       TD.ShowQty,"
StrSQL = StrSQL & "       LineDisc = ISNULL(TD.ItemDiscount,0) + ISNULL(TD.TotalDiscountPerLine,0),"
StrSQL = StrSQL & "       VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END,"

If PriceWithVAT Then
    ' TD.showPrice ????? ???? ??????? -> ?????? ???? ?????? ???? ?????
    StrSQL = StrSQL & "       BasePrice_Ex = ROUND("
    StrSQL = StrSQL & "           TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0)"
    StrSQL = StrSQL & "       , 6),"
    ' ?????? ????? ??? ????? (Ex-VAT) = Base * Qty
    StrSQL = StrSQL & "       LineEx_NoDisc = ROUND("
    StrSQL = StrSQL & "           (TD.showPrice * TD.ShowQty)"
    StrSQL = StrSQL & "             / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0)"
    StrSQL = StrSQL & "       , 6)"
Else
    ' TD.showPrice ????? ???? ?????
    StrSQL = StrSQL & "       BasePrice_Ex = ROUND(TD.showPrice, 6),"
    StrSQL = StrSQL & "       LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6)"
End If

StrSQL = StrSQL & "  FROM dbo.Transactions AS T"
StrSQL = StrSQL & "  INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & "  INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID"
StrSQL = StrSQL & "  WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
If SystemOptions.CanUploadZakatOpt Then
ConectionFirst:     Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText

' ------- ????? ?????????? -------
Dim SumLineExNoDisc As Double, SumLineDisc As Double
Dim SumLineExAfterDisc As Double, SumVAT As Double, SumInc As Double

Do Until rs.EOF
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")

    Dim ItemName As String:   ItemName = rs!ItemName & ""
    Dim Qty As Double:        Qty = val(rs!ShowQty & "")
    Dim basePrice As Double:  basePrice = val(rs!BasePrice_Ex & "")   ' ??? ?????? ??? ????? ????? ?????
    Dim lineDisc As Double:   lineDisc = Round(val(rs!lineDisc & ""), 6)
    Dim vatRate As Double:    vatRate = Round(val(rs!vatRate & ""), 6) ' ????? %

    ' ????? ?????? (Ex-VAT) ??? ?????
    invline.Price.EncludingVat = False
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = ItemName

    ' ????? ?????
    If vatRate > 0 And val(rs!ItemWithOutVAT & "") = 0 Then
        invline.Item.classifiedTaxCategory.ID = "S"
        invline.Item.classifiedTaxCategory.Percent = vatRate
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = vatRate
    Else
        invline.Item.classifiedTaxCategory.ID = "E"
        invline.Item.classifiedTaxCategory.Percent = 0
        invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E"
        invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
    End If

    ' ??? ????? ?? AllowanceCharge (????? ????? ???????)
    If lineDisc > 0# Then
        Dim lac As Object
        Set lac = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        lac.ChargeIndicator = False
        lac.Amount = Format(lineDisc, "0.######")
        lac.AllowanceChargeReason = "Line discount"
        lac.taxCategory.ID = invline.TaxTotal.TaxSubtotal.taxCategory.ID
        lac.taxCategory.Percent = invline.TaxTotal.TaxSubtotal.taxCategory.Percent
        invline.allowanceCharges.Add lac
    End If

    ' ?????? ?????:
    ' LineEx_NoDisc ????? ?? SQL (Ex-VAT)
    Dim lineExNoDisc As Double: lineExNoDisc = val(rs!LineEx_NoDisc & "")
    Dim lineExAfterDisc As Double: lineExAfterDisc = Round(lineExNoDisc - lineDisc, 6)
    If lineExAfterDisc < 0 Then lineExAfterDisc = 0

    Dim LineVat As Double
    If vatRate > 0 Then
        LineVat = Round(lineExAfterDisc * (vatRate / 100#), 6)
    Else
        LineVat = 0
    End If
    Dim lineInc As Double: lineInc = Round(lineExAfterDisc + LineVat, 6)

    ' ?????
    SumLineExNoDisc = SumLineExNoDisc + lineExNoDisc
    SumLineDisc = SumLineDisc + lineDisc
    SumLineExAfterDisc = SumLineExAfterDisc + lineExAfterDisc
    SumVAT = SumVAT + LineVat
    SumInc = SumInc + lineInc

    inv.InvoiceLines.Add invline
    rs.MoveNext
Loop
rs.Close

' ------- ????? ?????????? UBL -------
' AllowanceTotalAmount = ????? ?????? ?????? (????? ?? 5074 ?? ??????)
inv.LegalMonetaryTotal.LineExtensionAmount = Format(SumLineExAfterDisc, "0.##") ' ???? ?????? Ex-VAT ??? ?????
inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(SumLineExAfterDisc, "0.##")
inv.TaxTotal.TaxAmount = Format(SumVAT, "0.##")
inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(SumInc, "0.##")
inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(SumLineDisc, "0.##")
inv.LegalMonetaryTotal.ChargeTotalAmount = Format(0, "0.##")
inv.LegalMonetaryTotal.PayableRoundingAmount = Format(0, "0.##")
inv.LegalMonetaryTotal.PayableAmount = Format(SumInc, "0.##")

' ------- ??????? ???????? (?? ?? ????) -------
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)
If res.IsValid = True Then
    ' ???? ??????? + ???? API ??? ?? ???? ??????
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    ' (??????? ??? API ??? ñ ???? ?? ???? ??????)
    GenerateInvoice_ZatcaV35 = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV35 = False
End If

End Function

' Helper »”Ìÿ
Private Function Nz(v, Optional dflt As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then
        Nz = dflt
    Else
        Nz = v
    End If
End Function

Private Function NzText(v) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function









Function generateInvoiceOld(Optional ByVal docType As Integer = 0, Optional ByVal mTableName As String = "transactions", Optional ByVal mFieldIDName As String = "Transaction_ID", Optional ByVal ActivityTypeIdInv As Integer = 0, Optional ByVal branch_idInv As Integer = 0)

 On Error Resume Next: Err.Clear
' Dim ubl As ZatcaUBLXMLGenerator.UBLXML
'    Set ubl = New ZatcaUBLXMLGenerator.UBLXML
'Dim inv As ZatcaUBLXMLGenerator.invoice
'    Set inv = New ZatcaUBLXMLGenerator.invoice
'    Dim res As ZatcaUBLXMLGenerator.Result
'    Set res = New ZatcaUBLXMLGenerator.Result
'    '
     Dim ubl As Object
      Dim SendingMode As Integer
      
      Dim mDiscountProject As Double
     Dim SecretKey As String
 Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
    Dim inv As Object
 Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
  Dim res As Object
 Set res = CreateObject("ZatcaUBLXMLGenerator.Result")
    
     If cschemeID = "" Then
    cschemeID = "CRN"
 End If
 
 
  If SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
        Set cn2 = POSConnection
        
    Else
        Set cn2 = Cn
       
    End If
'    »Ì«‰«  «·›« Ê—…
'    QRCode = ""

If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

'ID = 12345
    inv.ID = ID ' "1230"
    If IssueDate = "" Then IssueDate = Date
    inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")           '"2021-01-05" 'IssueDate
    If IssueTim = "" Then IssueTim = Now
    inv.IssueTime = Format(IssueTim, "HH:MM:SS")   '"09:32:40"  ' Format ( IssueTime,"HH:MM:SS")
     ' ›« Ê—… ÷—Ì»Ì… «Ê „»”ÿ… 388
      ' «‘⁄«— „œÌ‰ 383
      ' 381 «‘⁄«— œ«∆‰
    inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
  inv.InvoiceTypeCode.Name = InvoiceTypeCodename
    'inv.invoiceTypeCode.Name based on format NNPNESB
    'NN 01 ··›« Ê—… «·÷—Ì»Ì…
    'NN 02 ··›« Ê—… «·÷—Ì»Ì… «·„»”ÿ…
    'P ›Ï Õ«·… ›« Ê—… ·ÿ—› À«·À ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'N ›Ï Õ«·… ›« Ê—… «”„Ì… ‰ﬂ » 1 ›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'E ›Ï Õ«·… ›« Ê—… ··’«œ—«  ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'S ›Ï Õ«·… ›« Ê—… „·Œ’… ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'B  ›Ï Õ«·… ›« Ê—… –« Ì… ‰ﬂ » 1
    'B ›Ï Õ«·… «‰ «·›« Ê—… ’«œ—« =1 ·« Ì„ﬂ‰ «‰  ﬂÊ‰ «·›« Ê—… –« Ì… =1
If DocumentCurrencyCode = "" Then DocumentCurrencyCode = "SAR"
If TaxCurrencyCode = "SR" Or TaxCurrencyCode = "" Then TaxCurrencyCode = "SAR"
If DocumentCurrencyCode = "SR" Or DocumentCurrencyCode = "" Then DocumentCurrencyCode = "SAR"

    inv.DocumentCurrencyCode = DocumentCurrencyCode
    inv.TaxCurrencyCode = TaxCurrencyCode
   'inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = "NWZlY2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTdlOQ=="
       'Ì „ «” œ⁄« «Œ— PIH
   'CalllastPIH
   inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)
 
 
    '›Ï Õ«·… «‘⁄«— œ«∆‰ «Ê „œÌ‰ ›ﬁÿ Â«‰ﬂ » —ﬁ„ «·›« Ê—… «··Ï «’œ—‰« «·«‘⁄«— ·ÌÂ«
 'inv.BillingReference.InvoiceDocumentReferenceID = val(InvoiceDocumentReferenceID)
 
    
 
    Dim InvoiceDocumentReference As Object
    Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
    InvoiceDocumentReference.ID = val(InvoiceDocumentReferenceID)
    inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    ' "?Invoice Number: 354; Invoice Issue Date: 2021-02-10?"
    ' ﬁÌ„… ⁄œ«œ «·›« Ê—…
    inv.AdditionalDocumentReferenceICV.UUID = uuidCounter ' 12345 val(AdditionalDocumentReferenceICVUUID) '„„ﬂ‰ ÌﬂÊ‰ —ﬁ„ «·›« Ê—… Ê·ﬂ‰ ·«»œ ÌﬂÊ‰ «—ﬁ«„ ›ﬁÿ
    ' ›Ï Õ«·… ›« Ê—… „»”ÿ… Ê›« Ê—… „·Œ’… Â«‰ﬂ »  «—ÌŒ «· ”·Ì„ Ê«Œ—  «—ÌŒ «· ”·Ì„

If ActualDeliveryDate = "" Then ActualDeliveryDate = IssueDate
If LatestDeliveryDate = "" Then LatestDeliveryDate = IssueDate
    inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")   ' ActualDeliveryDate ' "2022-10-22"
  If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")   ' LatestDeliveryDate '"2022-10-23"
    End If
   ' »Ì«‰«  «·œ›⁄
   ' «ﬂÊ«œ „⁄Ì‰…
   ' «Œ Ì«—Ï ﬂÊœ «·œ›⁄
       Dim paymentmean As Object
     Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
    paymentmean.PaymentMeansCode = PaymentMeansCode ' ????? ????? ????
     
    If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then '«‘⁄«— „œÌ‰ «Ê œ«∆‰
   paymentmean.InstructionNote = "Returned items" '›Ï Õ«·… «‘⁄«— «·œ«∆‰ Ê«·„œÌ‰ ·«»œ „‰ «œŒ«·Â«
   End If
      
    paymentmean.PayeeFinancialAccount.ID = PayeeFinancialAccount ' ?? ???? ??? ????? ?? ???? ????? ??? ???? ????? IBAN
    paymentmean.PayeeFinancialAccount.paymentnote = paymentnote ' ?? ???? ??? ????? ?? ???? ?????
    inv.PaymentMeans.Add paymentmean
    
     

 

 
    inv.SupplierParty.PartyIdentification.ID = cIdentificationID
    If Identificationid = "" Then
        'cschemeID ="OTH"
    End If
    
    inv.SupplierParty.PartyIdentification.schemeID = cschemeID
    inv.SupplierParty.PostalAddress.StreetName = cStreetName
    inv.SupplierParty.PostalAddress.AdditionalStreetName = cAdditionalStreetName ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.BuildingNumber = cBuildingNumber
    inv.SupplierParty.PostalAddress.PlotIdentification = cPlotIdentification ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.CityName = cCityName
    inv.SupplierParty.PostalAddress.PostalZone = cPostalZone
    inv.SupplierParty.PostalAddress.CountrySubentity = cCountrySubentity ' «Œ Ì«—Ì
    inv.SupplierParty.PostalAddress.CitySubdivisionName = cCitySubdivisionName
    inv.SupplierParty.PostalAddress.Country.IdentificationCode = "SA" 'cIdentificationCode
    inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim(Companynameforvat)
    inv.SupplierParty.PartyTaxScheme.CompanyID = cCompanyID

If mId(InvoiceTypeCodename, 1, 2) = "01" Then
    '»Ì«‰«  «·⁄„Ì·
    If Id700 <> "" Then
        inv.CustomerParty.PartyIdentification.ID = Id700  'Company_Comment
    Else
        
       ' If Identificationid = "" Then Identificationid = "N/A"
        
           
    
        inv.CustomerParty.PartyIdentification.ID = Identificationid 'Company_Comment
    End If
    
     If CompanyID = "" Then
                CompanyID = "300000000000003"
                schemeID = "OTH"
            End If
    
    inv.CustomerParty.PartyIdentification.schemeID = schemeID '
    If Identificationid = "N/A" Or Identificationid = "" Then
        StreetName = "N/A"
        AdditionalStreetName = "N/A"
        BuildingNumber = "N/A"
        CityName = "N/A"
        PostalZone = "00000"
        CountrySubentity = "N/A"
        CitySubdivisionName = "N/A"
    End If
    If IdentificationCode = "" Then IdentificationCode = "SA"
    
    inv.CustomerParty.PostalAddress.StreetName = StreetName
    inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName ' «Œ Ì«—Ì
    inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
    inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification ' «Œ Ì«—Ì
    
    inv.CustomerParty.PostalAddress.CityName = CityName
    inv.CustomerParty.PostalAddress.PostalZone = PostalZone
    inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity '«”„ «·Ê·«Ì… «Œ Ì«—Ï
    inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim(RegistrationName) 'Company_Name_Eng
    If Id700 <> "" Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = "N/A" 'VATRegNo
        inv.CustomerParty.PartyTaxScheme.CompanyID = "" 'VATRegNo
    Else
        'If CompanyID = "" Then CompanyID = "N/A"
        inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID 'VATRegNo
        
    End If
End If




If mId(InvoiceTypeCodename, 1, 2) = "02" And (mTableName = "tblEInvoice" Or mTableName = "PropertyDueBatchDetail") Then
    '»Ì«‰«  «·⁄„Ì·
        If CompanyID = "" Then
                CompanyID = "300000000000003"
                schemeID = "OTH"
            End If
    If Identificationid = "0" Then Identificationid = ""
    If Id700 <> "" Then
        inv.CustomerParty.PartyIdentification.ID = Id700  'Company_Comment
    Else
        inv.CustomerParty.PartyIdentification.ID = Identificationid 'Company_Comment
    End If
    inv.CustomerParty.PartyIdentification.schemeID = schemeID '
    inv.CustomerParty.PostalAddress.StreetName = StreetName
    inv.CustomerParty.PostalAddress.AdditionalStreetName = AdditionalStreetName ' «Œ Ì«—Ì
    inv.CustomerParty.PostalAddress.BuildingNumber = BuildingNumber
    inv.CustomerParty.PostalAddress.PlotIdentification = PlotIdentification ' «Œ Ì«—Ì
    
    inv.CustomerParty.PostalAddress.CityName = CityName
    inv.CustomerParty.PostalAddress.PostalZone = PostalZone
    inv.CustomerParty.PostalAddress.CountrySubentity = CountrySubentity '«”„ «·Ê·«Ì… «Œ Ì«—Ï
    inv.CustomerParty.PostalAddress.CitySubdivisionName = CitySubdivisionName
    If IdentificationCode = "" Then IdentificationCode = "SA"
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = IdentificationCode
    inv.CustomerParty.PartyLegalEntity.RegistrationName = RegistrationName 'Company_Name_Eng
    If Id700 <> "" Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = "N/A" 'VATRegNo
        inv.CustomerParty.PartyTaxScheme.CompanyID = "" 'VATRegNo
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = CompanyID 'VATRegNo
    End If
End If


    ' Â‰« ‰ﬁœ— ‰⁄„· loop Ê‰·› ⁄·Ï «·ﬂÊœ ÊÌ√„« ‰⁄„·Â Œ’„ «Ê «÷«›…
    ' »Ì«‰«  «·Œ’„ ⁄·Ï „” ÊÏ «·›« Ê—…
    Dim allowancecharge As Object
    ' ﬂÊœ «·Œ’„
    Set allowancecharge = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
'allowancechargeAmount = 0 '   0800000
 
    allowancecharge.Amount = allowancechargeAmount ' Format(allowancechargeAmount, ".##")     ' ﬁÌ„… «·Œ’„
  ' allowancecharge.MultiplierFactorNumeric = 10
  ' allowancecharge.BaseAmount = 100

    allowancecharge.AllowanceChargeReason = "Discount" ' ”»» «·Œ’„ ⁄·Ï „” ÊÏ «·›« Ê—…
    If chkTaxExempt = 1 Or TaxCategoryPercent = 0 Then
        allowancecharge.taxCategory.ID = "E" ' ﬂÊœ «·÷—Ì»…
        allowancecharge.taxCategory.Percent = 0# ' ‰”»… «·÷—Ì»…
    Else
    
        allowancecharge.taxCategory.ID = "S" ' ﬂÊœ «·÷—Ì»…
        allowancecharge.taxCategory.Percent = 15# ' ‰”»… «·÷—Ì»…
    End If
    inv.allowanceCharges.Add allowancecharge

    ' «Ã„«·Ï «·„»·€ »⁄œ «·÷—Ì»… - «·„»·€ «·„œ›Ê⁄
   inv.LegalMonetaryTotal.PayableAmount = (PayableAmount)
  ' «·„»·€ «·„œ›Ê⁄
     If val(InvoiceTypeCodeID) = 388 Then
   inv.LegalMonetaryTotal.PrepaidAmount = (PrepaidAmount)
   End If
   '********************************************************************************

   '**************************************************************************
    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    

     Dim rs As New ADODB.Recordset
     Dim StrSQL As String
 If docType = 0 Then
 
 Dim Percetage As Double

StrSQL = "SELECT dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, TblItems.ItemWithOutVAT, "
StrSQL = StrSQL & "dbo.Transaction_Details.ShowQty, PerforValue = 0,  "
StrSQL = StrSQL & "dbo.Transaction_Details.ItemDiscount AS discountvalue, Transactions.Trans_Discount as discountvalue2,"

' «· Õﬁﬁ „„« ≈–« ﬂ«‰ «·”⁄— Ì‘„· «·÷—Ì»…
If SystemOptions.PriceWithVAT Then
    PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage
    
    ' Õ”«» «·”⁄— »œÊ‰ «·÷—Ì»… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100)) "
    StrSQL = StrSQL & "ELSE dbo.Transaction_Details.showPrice "
    StrSQL = StrSQL & "END, 2) AS showPrice, "

    ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent, "

    ' ‰”»… «·÷—Ì»… «·„√ŒÊ–… „‰ «·„ €Ì— Ê·Ì” „‰ «·ÃœÊ·
    'StrSQL = StrSQL & " CAST(" & Percetage & " AS DECIMAL(18, 2)) AS vatPercent, "

    '  ÕœÌœ „« ≈–« ﬂ«‰ «·„‰ Ã Œ«÷⁄« ··÷—Ì»… √„ ·«
    StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN TblItems.ItemWithOutVAT = 0 THEN " & Percetage
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "
Else
    ' ›Ì Õ«·… «·”⁄— €Ì— ‘«„· «·÷—Ì»…
    StrSQL = StrSQL & "dbo.Transaction_Details.showPrice , "
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100), 2) AS VATAmount, "
    StrSQL = StrSQL & " ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent, dbo.Transaction_Details.Vatyo "
End If

'  ﬂ„·… «·«” ⁄·«„ »—»ÿ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.Transactions "
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
StrSQL = StrSQL & " WHERE (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"




'Dim Percetage As Double



StrSQL = "SELECT dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, TblItems.ItemWithOutVAT, "
StrSQL = StrSQL & "dbo.Transaction_Details.ShowQty, PerforValue = 0,  "
StrSQL = StrSQL & "dbo.Transaction_Details.ItemDiscount + dbo.Transaction_Details.TotalDiscountPerLine AS discountvalue2, "
StrSQL = StrSQL & "Transactions.Trans_Discount AS discountvalue, Transactions.Trans_DiscountType, "

' «· Õﬁﬁ „„« ≈–« ﬂ«‰ «·”⁄— Ì‘„· «·÷—Ì»…
If SystemOptions.PriceWithVAT Then
    PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage

    ' Õ”«» ≈Ã„«·Ì «·”⁄— ﬁ»· «·Œ’„
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty, 2) AS totalPrice, "

    ' Õ”«» ≈Ã„«·Ì «·”⁄— »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0), 2) AS totalPriceAfterDiscount, "

    ' «” Œ—«Ã «·”⁄— »œÊ‰ «·÷—Ì»… „‰ «·≈Ã„«·Ì »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "/ (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100)), 2) / dbo.Transaction_Details.ShowQty AS showPrice, "

    ' Õ”«» «·÷—Ì»… »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "- (((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "/ (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))), 2) AS vatPercent, "
    
    
      ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN IsNull(TblItems.ItemWithOutVAT,0) = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST(" & Percetage & " AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent2, "
    
      StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN  IsNull(TblItems.ItemWithOutVAT,0) = 0 THEN " & Percetage
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "


Else
    ' ≈–« ﬂ«‰ «·”⁄— €Ì— ‘«„· «·÷—Ì»…
   ' StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice, 2) AS showPrice, "


 StrSQL = StrSQL & "ROUND(((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0))  , 2) / dbo.Transaction_Details.ShowQty AS showPrice,"
    
    
    ' Õ”«» ≈Ã„«·Ì «·”⁄— ﬁ»· «·Œ’„
    StrSQL = StrSQL & "ROUND(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty, 2) AS totalPrice, "

    ' Õ”«» ≈Ã„«·Ì «·”⁄— »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0), 2) AS totalPriceAfterDiscount, "

    ' Õ”«» «·÷—Ì»… »⁄œ «·Œ’„
    StrSQL = StrSQL & "ROUND((dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty - IsNull(dbo.Transaction_Details.TotalDiscountPerLine,0)) "
    StrSQL = StrSQL & "* (CAST(dbo.Transaction_Details.Vatyo AS DECIMAL(18, 2)) / 100), 2) AS vatPercent, "
    
    
      ' Õ”«» «·÷—Ì»… «·„÷«›… »«” Œœ«„ ‰”»… «·÷—Ì»… „‰ «·„ €Ì— Percetage
    StrSQL = StrSQL & "ROUND(CASE "
    StrSQL = StrSQL & "WHEN IsNull(ItemWithOutVAT,0) = 0 THEN dbo.Transaction_Details.showPrice - (dbo.Transaction_Details.showPrice / (1 + (CAST( dbo.Transaction_Details.Vatyo AS DECIMAL(18, 2)) / 100))) "
    StrSQL = StrSQL & "ELSE 0 "
    StrSQL = StrSQL & "END, 2) AS vatPercent2, "
    
      StrSQL = StrSQL & "CASE "
    StrSQL = StrSQL & "WHEN IsNull(ItemWithOutVAT,0) = 0 THEN Vatyo"
    StrSQL = StrSQL & " ELSE 0 "
    StrSQL = StrSQL & "END AS Vatyo "

End If

' ‰”»… «·÷—Ì»…


'  ÕœÌœ „« ≈–« ﬂ«‰ «·„‰ Ã Œ«÷⁄« ··÷—Ì»… √„ ·«


'  ﬂ„·… «·«” ⁄·«„ »—»ÿ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.Transactions "
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
StrSQL = StrSQL & " WHERE (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"



'    StrSQL = "SELECT     dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,TblItems.ItemWithOutVAT, dbo.Transaction_Details.ShowQty,PerforValue = 0, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemDiscount as discountvalue, "
'      If SystemOptions.PriceWithVAT Then
'        PercentgValueAddedAccount_Transec Date, 21, 0, , Percetage
'        StrSQL = StrSQL & "                     ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent,   "
'
'        StrSQL = StrSQL & "  CASE"
'        StrSQL = StrSQL & "  WHEN TblItems.ItemWithOutVAT = 0 THEN " & Percetage
'        StrSQL = StrSQL & "  ELSE 0"
'        StrSQL = StrSQL & "  END As Vatyo"
'
'
'    Else
'        StrSQL = StrSQL & "                     ROUND(dbo.Transaction_Details.Vat, 2) AS vatPercent, dbo.Transaction_Details.Vatyo"
'    End If
'
'    StrSQL = StrSQL & "    FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL & "                          dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    StrSQL = StrSQL & "                          dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
'    StrSQL = StrSQL & "     Where (dbo.transactions.Transaction_ID = " & Transaction_ID & ")"
ElseIf docType = 1 Then

    

     StrSQL = " SELECT     project_bill_details.item as   ItemName, project_bill_details.item as ItemNamee, ItemWithOutVAT = 0, 1 as ShowQty,project_billl.PerforValue,"
     
     
     
   StrSQL = StrSQL & "               showPrice = IsNull(dbo.project_bill_details.totEx,0) - (IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0))"
   StrSQL = StrSQL & "               ,(IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0)) AS discountvalue"
   StrSQL = StrSQL & "               ,ROUND(IsNull(dbo.project_bill_details.LineVat,0), 2) AS vatPercent22"
   StrSQL = StrSQL & "               ,15 AS Vatyo"
   StrSQL = StrSQL & "               ,ROUND(( dbo.project_bill_details.totEx - (IsNull(project_billl.discount,0) + IsNull(project_billl.DiscountGMater,0) + IsNull(project_billl.advancedPayment,0)) ) * .15, 2) AS vatPercent"
     
     'StrSQL = StrSQL & "                         showPrice  = dbo.project_bill_details.totEx  - (IsNull(project_billl.discount,0) + isnull(project_billl.DiscountGMater,0) + isnull(project_billl.advancedPayment,0)) , "
'     StrSQL = StrSQL & "                         showPrice  = dbo.project_bill_details.totEx  ROUND((dbo.project_bill_details.totEx - (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment))   , "
'     StrSQL = StrSQL & "                         (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment) as discountvalue,"
'     StrSQL = StrSQL & "                         ROUND(dbo.project_bill_details.LineVat, 2) AS vatPercent22, 15 as Vatyo,"
'    ' StrSQL = StrSQL & "                         ROUND((dbo.project_bill_details.totEx -( project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment)) * .15,2) AS vatPercent"
'     StrSQL = StrSQL & "                         ROUND((dbo.project_bill_details.totEx ) * .15,2) AS vatPercent,"
'
'     StrSQL = StrSQL & "                    ROUND((dbo.project_bill_details.totEx - (project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment)) * 0.15, 2) As VatPercent"
     StrSQL = StrSQL & "             FROM         dbo.project_billl INNER JOIN"
     StrSQL = StrSQL & "     dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id"
     
     StrSQL = StrSQL & "     Where (dbo.project_billl.ID = " & Transaction_ID & ")"
     
     


' »‰«¡ «·«” ⁄·«„ «·√”«”Ì


' »‰«¡ «·«” ⁄·«„ «·√”«”Ì
StrSQL = " SELECT project_bill_details.item AS ItemName, "
StrSQL = StrSQL & " project_bill_details.item AS ItemNamee, "
StrSQL = StrSQL & " 0 AS ItemWithOutVAT, "
StrSQL = StrSQL & " 1 AS ShowQty, "
StrSQL = StrSQL & " project_billl.PerforValue, "

'  ⁄œÌ· Õﬁ· «·”⁄— »‰«¡ ⁄·Ï ﬁÌ„… SystemOptions.SuppCreat4Acc
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ ÿ—Õ «·Œ’„ «·„Ê“⁄
    'project_billl.Discount

   

    
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " CASE "
    StrSQL = StrSQL & " WHEN project_billl.discount1ID = 1 THEN "
    StrSQL = StrSQL & " dbo.project_bill_details.totEx * (IsNull(project_billl.discount1value, 0) / 100) "
    StrSQL = StrSQL & " ELSE "
    StrSQL = StrSQL & " IsNull(project_billl.discount1value, 0) "
    StrSQL = StrSQL & " END, "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ «” Œœ«„ «·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

' Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·”⁄— »⁄œ «·Œ’„
StrSQL = StrSQL & " vatPercent = ROUND((dbo.project_bill_details.totEx - "
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·Œ’„ «·„Ê“⁄
    StrSQL = StrSQL & " CASE "
    StrSQL = StrSQL & " WHEN project_billl.discount1ID = 1 THEN "
    StrSQL = StrSQL & " dbo.project_bill_details.totEx * (IsNull(project_billl.discount1value, 0) / 100) "
    StrSQL = StrSQL & " ELSE "
    StrSQL = StrSQL & " IsNull(project_billl.discount1value, 0) "
    StrSQL = StrSQL & " END "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ Õ”«» «·÷—Ì»… »«·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)) "
End If
StrSQL = StrSQL & ") * 0.15, 2), "

' ≈÷«›… »«ﬁÌ «·ÕﬁÊ·
StrSQL = StrSQL & " discountvalue = (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
StrSQL = StrSQL & " ROUND(IsNull(dbo.project_bill_details.LineVat, 0), 2) AS vatPercent22, "
StrSQL = StrSQL & " 15 AS Vatyo "

' «·—»ÿ »Ì‰ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.project_billl INNER JOIN "
StrSQL = StrSQL & " dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id "

' «·‘—ÿ
StrSQL = StrSQL & " WHERE (dbo.project_billl.ID = " & Transaction_ID & ")"


 
 ' »‰«¡ «·«” ⁄·«„ «·√”«”Ì
StrSQL = " SELECT project_bill_details.item AS ItemName, "
StrSQL = StrSQL & " project_bill_details.item AS ItemNamee, "
StrSQL = StrSQL & " 0 AS ItemWithOutVAT, "
StrSQL = StrSQL & " 1 AS ShowQty, "
StrSQL = StrSQL & " project_billl.PerforValue, "

'  ⁄œÌ· Õﬁ· «·”⁄— »‰«¡ ⁄·Ï ﬁÌ„… SystemOptions.SuppCreat4Acc
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „  Ê“Ì⁄ «·Œ’„ „‰ ›Ì·œ Discount ›ﬁÿ
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.Discount, 0) * "
    StrSQL = StrSQL & " (dbo.project_bill_details.totEx / "
    StrSQL = StrSQL & " (SELECT SUM(dbo.project_bill_details.totEx) "
    StrSQL = StrSQL & " FROM dbo.project_bill_details "
    StrSQL = StrSQL & " WHERE dbo.project_bill_details.bill_id = " & Transaction_ID & "))), "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ «” Œœ«„ «·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " showPrice = dbo.project_bill_details.totEx - "
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

' Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·”⁄— »⁄œ «·Œ’„
StrSQL = StrSQL & " vatPercent = ROUND((dbo.project_bill_details.totEx - "
If SystemOptions.SuppCreat4Acc = False Then
    ' ≈–« ﬂ«‰  «·ﬁÌ„… False° Ì „ Õ”«» «·÷—Ì»… »‰«¡ ⁄·Ï «·Œ’„ «·„Ê“⁄
    StrSQL = StrSQL & " (IsNull(project_billl.Discount, 0) * "
    StrSQL = StrSQL & " (dbo.project_bill_details.totEx / "
    StrSQL = StrSQL & " (SELECT SUM(dbo.project_bill_details.totEx) "
    StrSQL = StrSQL & " FROM dbo.project_bill_details "
    StrSQL = StrSQL & " WHERE dbo.project_bill_details.bill_id = " & Transaction_ID & "))) "
Else
    ' ≈–« ﬂ«‰  «·ﬁÌ„… True° Ì „ Õ”«» «·÷—Ì»… »«·ÿ—Ìﬁ… «·ﬁœÌ„…
    StrSQL = StrSQL & " (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)) "
End If
StrSQL = StrSQL & ") * 0.15, 2), "

' ≈÷«›… »«ﬁÌ «·ÕﬁÊ·
If SystemOptions.SuppCreat4Acc = False Then
    StrSQL = StrSQL & " discountvalue = IsNull(project_billl.discount, 0), "
Else
    StrSQL = StrSQL & " discountvalue = (IsNull(project_billl.discount, 0) + IsNull(project_billl.DiscountGMater, 0) + IsNull(project_billl.advancedPayment, 0)), "
End If

StrSQL = StrSQL & " ROUND(IsNull(dbo.project_bill_details.LineVat, 0), 2) AS vatPercent22, "
StrSQL = StrSQL & " 15 AS Vatyo "

' «·—»ÿ »Ì‰ «·Ãœ«Ê·
StrSQL = StrSQL & " FROM dbo.project_billl INNER JOIN "
StrSQL = StrSQL & " dbo.project_bill_details ON dbo.project_billl.id = dbo.project_bill_details.bill_id "

' «·‘—ÿ
StrSQL = StrSQL & " WHERE (dbo.project_billl.ID = " & Transaction_ID & ")"


ElseIf docType = 3 Then

    

     StrSQL = " SELECT    'Notes'   ItemName, 'Notes' as ItemNamee, ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, Notes.Note_Value as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "                         ROUND(dbo.Notes.VAT, 2) AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "             FROM         "
     StrSQL = StrSQL & "     dbo.Notes "
     
     StrSQL = StrSQL & "     Where (dbo.Notes.NoteID = " & Transaction_ID & ")"

ElseIf docType = 4 Then
     StrSQL = " SELECT   top 1  ExpensesRemark   ItemName, ExpensesRemark as ItemNamee,ItemWithOutVAT = 0, Count as ShowQty,0 as PerforValue, Notes.price as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.Notes.Note_Value  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.notes"
      StrSQL = StrSQL & " Where IsNull(notes.CusID, 0) = 0 And notes_all = " & Transaction_ID & ""
      
      
ElseIf docType = 10 Then

    
    StrSQL = "SELECT     ItemName,ItemWithOutVAT = 0, Qty ShowQty,PerforValue = 0, Price showPrice, discountvalue = 0, "
    StrSQL = StrSQL & "                      VatValue as Vatyo ,"
    StrSQL = StrSQL & "                vatPercent =   IsNull(VatValue,0) / (Price *Qty ) * 100  , "
    StrSQL = StrSQL & "                Vatyo =   IsNull(VatValue,0) / (Price *Qty ) * 100   "
    
    StrSQL = StrSQL & "    FROM         tblEInvoice2"
    StrSQL = StrSQL & "     Where (dbo.tblEInvoice2.InvoiceID = '" & order_no & "')"
ElseIf docType = 5 Then
    StrSQL = " SELECT   top 1  IsNull(Remarks,'Note')   ItemName, Remarks as ItemNamee,ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, TblHandWages.TotalBvat as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.TblHandWages.TotalBvat  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.TblHandWages"
      StrSQL = StrSQL & " Where id = " & Transaction_ID & ""
ElseIf docType = 6 Then
    StrSQL = " SELECT   top 1  IsNull(Remarks,'Note')   ItemName, Remarks as ItemNamee,ItemWithOutVAT = 0, 1 as ShowQty,0 as PerforValue, TblHandWages.TotalBvat as showPrice, 0 as discountvalue,"
     StrSQL = StrSQL & "             ROUND((dbo.TblHandWages.TotalBvat  * 0.15) , 2)  AS vatPercent, 15 as Vatyo"
     StrSQL = StrSQL & "              From"
      StrSQL = StrSQL & " dbo.TblHandWages"
      
      
    StrSQL = " SELECT  top 1 "
    StrSQL = StrSQL & "tblContractInsAllocationsDetails.ID"
    StrSQL = StrSQL & ",TblAqar.aqarname + ' ' + TblAqarDetai.unitno  as ItemName ,ItemWithOutVAT = 0"
    StrSQL = StrSQL & ",1 as ShowQty ,0 as PerforValue,tblContractInsAllocationsDetails.installValue - tblContractInsAllocationsDetails.VATValue as showPrice"
    StrSQL = StrSQL & ", 0 as discountvalue,tblContractInsAllocationsDetails.ContNo"
    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.VATValue vatPercent,TblContract.FATYou as Vatyo "
'    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.HNoteSerial1"
'    StrSQL = StrSQL & ",tblContractInsAllocationsDetails.installValue"
    StrSQL = StrSQL & "From TblAqarDetai "
    StrSQL = StrSQL & "INNER JOIN TblAqar "
    StrSQL = StrSQL & "    ON TblAqarDetai.Aqarid = TblAqar.Aqarid "
    StrSQL = StrSQL & "RIGHT OUTER JOIN TblContractInstallments "
    StrSQL = StrSQL & "INNER JOIN TblContract "
    StrSQL = StrSQL & "    ON TblContractInstallments.ContNo = TblContract.ContNo "
    StrSQL = StrSQL & "INNER JOIN tblContractInsAllocationsDetails "
    StrSQL = StrSQL & "    ON TblContractInstallments.InstallNo = tblContractInsAllocationsDetails.InstallNo "
    StrSQL = StrSQL & "        AND TblContract.ContNo = tblContractInsAllocationsDetails.ContNo "
    StrSQL = StrSQL & "    ON TblAqar.Aqarid = TblContract.Iqar "
    StrSQL = StrSQL & "        AND TblAqarDetai.Id = TblContract.UnitNo "


    StrSQL = StrSQL & " Where tblContractInsAllocationsDetails.id = " & Transaction_ID & ""
    
    
    If SystemOptions.CanUploadZakatOpt Then
'                StrSQL = " SELECT TOP 1 "
        StrSQL = ""
        
        ' œ«·… ’€Ì—… ÃÊ¯« SQL · ÕÊÌ· 0.15 ? 15 œÌ‰«„ÌﬂÌ«
        ' (Â  ﬂ » inline ﬂ‹ CASE)
        ' VAT% ··»‰œ: ·Ê ÷—«∆» «·»‰œ > 0 ‰«Œœ ‰”»… «·⁄ﬁœ° Ê≈·« 0
        
        ' ========= ≈ÌÃ«— =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ≈ÌÃ«—' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ≈ÌÃ«—' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchRentValueTaxes,0)=0 AND ISNULL(B.BatchRentValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchRentValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchRentValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchRentValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchRentValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= ﬂÂ—»«¡ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ﬂÂ—»«¡' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - ﬂÂ—»«¡' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchElectricityValueTaxes,0)=0 AND ISNULL(B.BatchElectricityValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchElectricityValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchElectricityValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchElectricityValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchElectricityValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= „Ì«Â =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - „Ì«Â' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - „Ì«Â' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchWaterValueTaxes,0)=0 AND ISNULL(B.BatchWaterValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchWaterValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchWaterValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchWaterValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchWaterValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= Œœ„«  =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - Œœ„« ' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - Œœ„« ' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchServicesValueTaxes,0)=0 AND ISNULL(B.BatchServicesValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchServicesValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchServicesValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchServicesValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchServicesValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= «·”¯Û⁄Ì («·⁄„Ê·…) =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - «·”⁄Ì' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - «·”⁄Ì' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchCommissionValueTaxes,0)=0 AND ISNULL(B.BatchCommissionValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchCommissionValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchCommissionValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchCommissionValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchCommissionValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' ========= €«“ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - €«“' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' - €«“' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchGasValueTaxes,0)=0 AND ISNULL(B.BatchGasValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchGasValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchGasValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchGasValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchGasValue,0) <> 0 "
        
        StrSQL = StrSQL & " UNION ALL "
        
        ' =========  √„Ì‰ =========
        StrSQL = StrSQL & " SELECT D.ID, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' -  √„Ì‰' AS ItemName, " & _
                          "        P.ArName + ' ' + PD.PropertyUnitNo + N' -  √„Ì‰' AS ItemNamee, " & _
                          "        CASE WHEN ISNULL(B.BatchInsuranceValueTaxes,0)=0 AND ISNULL(B.BatchInsuranceValue,0)<>0 THEN 1 ELSE 0 END AS ItemWithOutVAT, " & _
                          "        1 AS ShowQty, 0 AS PerforValue, " & _
                          "        CAST(ROUND(ISNULL(B.BatchInsuranceValue,0), 2) AS DECIMAL(18,2)) AS showPrice, " & _
                          "        0 AS discountvalue, C.Id AS ContNo, " & _
                          "        CASE WHEN ISNULL(B.BatchInsuranceValueTaxes,0)>0 " & _
                          "             THEN CAST(CASE WHEN C.VATPercentage>1 THEN C.VATPercentage ELSE C.VATPercentage*100 END AS DECIMAL(5,2)) " & _
                          "             ELSE 0 END AS Vatyo, " & _
                          "        CAST(ROUND(ISNULL(B.BatchInsuranceValueTaxes,0), 2) AS DECIMAL(18,2)) AS vatPercent " & _
                          " FROM PropertyDueBatchDetail D " & _
                          " INNER JOIN PropertyContractBatch B ON B.Id = D.PropertyContractBatchId " & _
                          " INNER JOIN PropertyContract C ON C.Id = B.MainDocId " & _
                          " INNER JOIN PropertyDetail PD ON C.PropertyUnitId = PD.Id " & _
                          " INNER JOIN Property P ON P.Id = C.PropertyId " & _
                          " WHERE D.ID = " & Transaction_ID & _
                          "   AND ISNULL(B.BatchInsuranceValue,0) <> 0 "

    
    
    End If
    
     
End If
Dim mValueWithOutVat As Double
Dim ISVAT As Boolean
Dim classifiedTaxCategory2 As Double

    Set rs = New ADODB.Recordset
    If SystemOptions.CanUploadZakatOpt Then
       ' ConectionFirst
     '   Set cn2 = POSConnection
        rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
    Else
        Set cn2 = Cn
        rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText
    End If
    PayableAmount = 0
    Totallinewithvat = 0
    Totallinewithoutvat = 0
    allowancechargeAmount = val(rs!discountvalue & "")
    If docType = 0 Then
     invline.Price.EncludingVat = SystemOptions.PriceWithVAT
   Else
        invline.Price.EncludingVat = False
    End If
   ISVAT = False
  For i = 1 To rs.RecordCount
       Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
        If docType = 10 Then
            invline.Price.PriceAmount = (IIf(IsNull(rs("showPrice").value), 0, (rs("showPrice").value))) '”⁄— «·„‰ Ã
        Else
            invline.Price.PriceAmount = Round(IIf(IsNull(rs("showPrice").value), 0, (rs("showPrice").value)), 2) '”⁄— «·„‰ Ã
        End If
        invline.InvoiceQuantity = IIf(IsNull(rs("ShowQty").value), 0, rs("ShowQty").value)
        
       ' discountvalue = IIf(IsNull(rs("discountvalue").value), 0, rs("discountvalue").value)
      '  linediscountvalue = IIf(IsNull(rs("discountvalue").value), 0, (rs("discountvalue").value))   '”⁄— «·„‰ Ã
        mDiscountProject = 0
        discountvalue = 0
    'mDiscountProject = IIf(IsNull(rs("PerforValue").value), 0, (rs("PerforValue").value))  'Œ’„ Õ”‰ «·«œ«¡ Õ«·… «·„‘«—Ì⁄ ÊÂÊ Ì „ Œ’„Â Ê·« ÌœŒ· ›Ï «·÷—Ì»…
        If linediscountvalue > 0 Then
            Debug.Print ID
        End If
        
            invline.Item.Name = IIf(IsNull(rs("ItemName").value), 0, rs("ItemName").value) ' «”„ «·„‰ Ã
            classifiedTaxCategory = Round(IIf(IsNull(rs("Vatyo").value), 0, rs("Vatyo").value))
            
            VatPercent = Round(IIf(IsNull(rs("vatPercent").value), 0, rs("vatPercent").value), 3)
            
            classifiedTaxCategory = Round(IIf(IsNull(rs("Vatyo").value), 0, rs("Vatyo").value))
VatPercent = Round(IIf(IsNull(rs("vatPercent").value), 0, rs("vatPercent").value), 3)

' Â‰« ﬂ«‰ ›Ì ·ÊÃÌﬂ »ÌÕÊ¯· VatPercent ·‹ 15 ·Ê ﬁ—Ì» „‰ 15
' œÂ Œÿ√ ·√‰Â VatPercent ⁄‰œ‰« "ﬁÌ„… ÷—Ì»…" „‘ "‰”»…"

'
'            If VatPercent < 15 And VatPercent > 13.99 Then
'                VatPercent = 15
'            ElseIf VatPercent > 15 And VatPercent < 16 Then
'                VatPercent = 15
'            End If
'            If VatPercent <> 15 Then
'                VatPercent = VatPercent
'            End If
             If VatPercent = 0 Then classifiedTaxCategory = 0 ': mValueWithOutVat = mValueWithOutVat + (val(invline.ShowPrice) * val(invline.InvoiceQuantity))
        If classifiedTaxCategory > 0 Then
            ISVAT = True
            classifiedTaxCategory2 = classifiedTaxCategory
            invline.Item.classifiedTaxCategory.ID = "S" 'Œ«÷⁄ ·÷—Ì»…
            invline.TaxTotal.TaxSubtotal.taxCategory.ID = "S" ' Œ«÷⁄ ··÷—Ì»…
            invline.Item.classifiedTaxCategory.Percent = classifiedTaxCategory '& "#"   ' ‰”»… «·÷—Ì»…
            invline.TaxTotal.TaxSubtotal.taxCategory.Percent = classifiedTaxCategory ' & "#"
            
        
    
        
        
        Else
    
    '    invline.taxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Not subject to VAT" '«·”»» ·Ê ﬂÊœ «·÷—Ì»… ﬂ«‰ O,E »„⁄‰Ï «‰Â €Ì— Œ«÷⁄ ··÷—Ì»…
    '    'invline.TaxTotal.TaxSubtotal.TaxCategory.TaxExemptionReason = "" '«·”»» ·Ê ﬂÊœ «·÷—Ì»… ﬂ«‰ E »„⁄‰Ï «‰Â „⁄›Ï „‰ ÷—Ì»… «·ﬁÌ„… «·„÷«›…
    '    invline.taxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "" ' ›Ï Õ«·… ﬂ«‰ ﬂÊœ «·÷—Ì»… Z ·«»œ „‰ «œŒ«· ﬂÊœ «⁄›«¡ «·÷—Ì»…
    
    '«ﬂÊ«œ «·«⁄›«¡« 
            If InvoiceTypeCodename = "0100100" Then  'export
                invline.Item.classifiedTaxCategory.ID = "Z" 'ﬂÊœ «·÷—Ì»…
                invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z" 'ﬂÊœ «·÷—Ì»…
                invline.Item.classifiedTaxCategory.Percent = 0
                
                invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-32" '   ﬂÊœ «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Export of goods" ' ”»» «·«⁄›«¡
        
            End If
            
            If InvoiceTypeCodename = "0100000" And docType2 <> 6 Then  'export
                invline.Item.classifiedTaxCategory.ID = "E" 'ﬂÊœ «·÷—Ì»…
                invline.Item.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29" 'ﬂÊœ «·÷—Ì»…
                invline.Item.taxCategory.TaxExemptionReason = "Financial services specified in Article 29 of the VAT Implementing Regulations" ' ”»» «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E" 'ﬂÊœ «·÷—Ì»…
                invline.Item.classifiedTaxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29" '   ﬂÊœ «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29 of the VAT Implementing Regulations" ' ”»» «·«⁄›«¡
                
            ElseIf InvoiceTypeCodename = "0100000" And docType2 = 6 Then   'export
            
                invline.Item.classifiedTaxCategory.ID = "O" 'ﬂÊœ «·÷—Ì»…
                invline.TaxTotal.TaxSubtotal.taxCategory.ID = "O" 'ﬂÊœ «·÷—Ì»…
                invline.Item.classifiedTaxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "" '   ﬂÊœ «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Out of Scope" ' ”»» «·«⁄›«¡

            End If
            
            
             If InvoiceTypeCodename = "0200000" And classifiedTaxCategory = 0 Then    'export
              invline.Item.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29" 'ﬂÊœ «·÷—Ì»…
                invline.Item.taxCategory.TaxExemptionReason = "Financial services specified in Article 29 of the VAT Implementing Regulations" ' ”»» «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.ID = "E" 'ﬂÊœ «·÷—Ì»…
                invline.Item.classifiedTaxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-29" '   ﬂÊœ «·«⁄›«¡
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Financial services specified in Article 29 of the VAT Implementing Regulations" ' ”»» «·«⁄›«¡
        
        
            End If
        End If
        If Round(IIf(IsNull(rs("showPrice").value), 0, (rs("showPrice").value)), 2) = 0 Then
                invline.Item.classifiedTaxCategory.ID = "Z"                          ' ﬂÊœ «·÷—Ì»… ’›—Ì… (Zero Rated)
                invline.Item.classifiedTaxCategory.Percent = 0                       ' «·‰”»… ’›—
                invline.Item.taxCategory.ID = "Z"                                    ' ’›—Ì… (·Ê «·Õﬁ· „ÿ·Ê»)
                invline.Item.taxCategory.Percent = 0                                 ' «·‰”»… ’›—
                invline.Item.taxCategory.TaxExemptionReasonCode = "VATEX-SA-00"     ' ﬂÊœ ≈⁄›«¡ «› —«÷Ì ··√’‰«› ’›—Ì… «·÷—Ì»… (·Ê «·Õﬁ· ≈Ã»«—Ì)
                invline.Item.taxCategory.TaxExemptionReason = "Free sample"          ' ”»» ≈⁄›«¡ ⁄«„ („À·«: Free sample √Ê "⁄Ì‰… „Ã«‰Ì…")
                invline.TaxTotal.TaxSubtotal.taxCategory.ID = "Z"                   ' ﬂÊœ «·÷—Ì»… ’›—Ì…
                invline.TaxTotal.TaxSubtotal.taxCategory.Percent = 0                ' «·‰”»… ’›—
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = "VATEX-SA-00" ' ‰›” ﬂÊœ «·≈⁄›«¡ «·«› —«÷Ì (·Ê „ÿ·Ê»)
                invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = "Free sample"     ' ”»» ≈⁄›«¡ ⁄«„ (·Ê „ÿ·Ê»)



        End If

    ' ·Ê Â ” Œœ„ «·”⁄— ‘«„· «·÷—Ì»… Ì»ﬁÏ EncludingVat=True
   
 
If docType = 10 Then
   invline.InvoiceQuantity = Format(invline.InvoiceQuantity, ".##########")
  invline.Price.PriceAmount = Format(invline.Price.PriceAmount, ".##########")
  discountvalue = Format(discountvalue, "0.##")
    lineTotal = ((invline.InvoiceQuantity * invline.Price.PriceAmount - discountvalue))
  ' lineTotal = ((invline.InvoiceQuantity * invline.Price.PriceAmount))
  lineTotal = Format(lineTotal, ".##########")

Else
   invline.InvoiceQuantity = Format(invline.InvoiceQuantity, ".##")
  invline.Price.PriceAmount = Format(invline.Price.PriceAmount, ".##")
  discountvalue = Format(discountvalue, "0.##")
    lineTotal = ((invline.InvoiceQuantity * invline.Price.PriceAmount - discountvalue))
  ' lineTotal = ((invline.InvoiceQuantity * invline.Price.PriceAmount))
  lineTotal = Format(lineTotal, ".##")
End If
  allowancechargeAmount = 0
  allowancechargeAmount = allowancechargeAmount + discountvalue
 
  If classifiedTaxCategory > 0 Then
'    If docType2 = 6 Then
'
        'Totallinewithvat = Totallinewithvat + (rs!VatPercent & "")
'    Else
'        Totallinewithvat = Totallinewithvat + lineTotal
'    End If
    Totallinewithvat = Totallinewithvat + lineTotal
    
 '„ƒﬁ « ·ÕÌ‰  Ã—» Â« ÃÌœ« ·«” «– ÿ·«·  „ ⁄„· ﬂÊ„‰  ›Ï «·”ÿ— »«·«⁄·Ï
  '  Totallinewithvat = (Totallinewithvat + lineTotal) * 1.15
  
  Else
    Totallinewithoutvat = Totallinewithoutvat + lineTotal
  End If
 
 'Totallinewithoutvat = lineTotal
 'Totallinewithvat = (lineTotal) * 1.15
     ' «‰ ÊÃœ ⁄·Ï „” ÊÏ ”⁄— «·„‰ Ã
 
'     If linediscountvalue > 0 And 1 = 0 Then
'     Dim invlineallowancecharge As ZatcaUBLXMLGenerator.allowancecharge
'    Set invlineallowancecharge = New ZatcaUBLXMLGenerator.allowancecharge
'    invlineallowancecharge.AllowanceChargeReasonCode = "Discount"
'    invlineallowancecharge.Amount = linediscountvalue
'    invline.allowanceCharges.Add invlineallowancecharge
 
    ' invline.Price.allowancecharge.AllowanceChargeReason = "Reason- commercial discount"
    ' invline.Price.allowancecharge.Amount = linediscountvalue

    ' «‰ ÊÃœ ⁄·Ï „” ÊÏ «· invoice line ﬂ·…
'    invline.allowancecharge.AllowanceChargeReason = "Discount"
'     invline.allowancecharge.Amount = linediscountvalue

'   End If






    inv.InvoiceLines.Add invline
    rs.MoveNext
Next i


'If ISVAT Then
'    classifiedTaxCategory = classifiedTaxCategory2
'    invline.taxTotal.TaxSubtotal.taxCategory.ID = "S" ' Œ«÷⁄ ··÷—Ì»…
'    invline.Item.classifiedTaxCategory.Percent = classifiedTaxCategory '& "#"   ' ‰”»… «·÷—Ì»…
'    invline.taxTotal.TaxSubtotal.taxCategory.Percent = classifiedTaxCategory ' & "#"
'End If
            
   'mDiscountProject
 '·ﬁœ  „ «Ìﬁ«› Â–« «·”ÿ— ·«‰‰« ﬁ„‰« » Ê“Ì⁄ «·Œ’„ ⁄·Ï „” ÊÏ «·”ÿ— Ê „ Œ’„Â „”»ﬁ«
 '  TotalVat = (Totallinewithvat - allowancechargeAmount) * classifiedTaxCategory / 100
   TotalVat = (Totallinewithvat) * classifiedTaxCategory / 100
  '›Ï ‰›” «·Õ«·… «·”«»ﬁ… Â‰⁄„· ﬂÊ„‰  ⁄·Ï «·”ÿ— «··Ï ›Êﬁ Ê‰Ã—» „ƒﬁ «
  
  
  '  Ã—»…
'  If docType2 = 10 Then
'    TotalVat = Vat
'  Else
'    TotalVat = (lineTotal) * classifiedTaxCategory / 100
'End If
   '  Ã—»…
   
   
'inv.ChargeTotalAmount = mDiscountProject
If docType = 10 Then
  TotalVat = val(Format(TotalVat, ".##########"))
Else
    TotalVat = val(Format(TotalVat, ".##"))

End If

  '·ﬁœ  „ «Ìﬁ«› Â–« «·”ÿ— ·«‰‰« ﬁ„‰« » Ê“Ì⁄ «·Œ’„ ⁄·Ï „” ÊÏ «·”ÿ— Ê „ Œ’„Â „”»ﬁ«
  'PayableAmount = Totallinewithvat + TotalVat + Totallinewithoutvat - mDiscountProject - allowancechargeAmount
  PayableAmount = Totallinewithvat + TotalVat + Totallinewithoutvat - mDiscountProject
  '›Ï ‰›” «·Õ«·… «·”«»ﬁ… Â‰⁄„· ﬂÊ„‰  ⁄·Ï «·”ÿ— «··Ï ›Êﬁ Ê‰Ã—» „ƒﬁ «
 ' PayableAmount = Totallinewithvat + TotalVat + Totallinewithoutvat - mDiscountProject - allowancechargeAmount
  
  'inv.LegalMonetaryTotal.ChargeTotalAmount = Format(mDiscountProject, ".##")
  If docType = 10 Then
        inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmount, ".##########")
        allowancecharge.Amount = Format(allowancechargeAmount, "0.##########")
    Else
        inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmount, ".##")
        allowancecharge.Amount = Format(allowancechargeAmount, "0.##")

    End If
        
        
'    If dev.value = True Then
'        SendingMode = 0
'    ElseIf sim.value = True Then
'        SendingMode = 1
'    ElseIf Live.value = True Then
'        SendingMode = 2
'    End If
If mTableName = "tblEInvoice" Then
    mFieldIDName = "ID"
End If

SendingMode = SystemOptions.SendingMode

'here you can pass csid data
'„„ﬂ‰ ‰—”· »Ì«‰«  «· ceetificate „‰ «·œ« «»Ì“
If SendingMode <> 0 Then
 inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem  ' "" 'from file cert.pem
 inv.CSIDInfo.Privatekey = SystemOptions.Privatekey 'from file key.pem
'Dim SecretKey As String
SecretKey = SystemOptions.SecretKey  'from secret.txt file
End If


'WaelZakah
Set res = ubl.GenerateInvoiceXML(inv, App.path)




If mId(InvoiceTypeCodename, 1, 2) = "02" Then 'simple invoice

If res.IsValid = True Then
' Â–… «·‰ «∆Ã Ì—ÃÏ Õ›ŸÂ« ›Ï ﬁ«⁄œ… «·»Ì«‰«  «·Œ«’… »«·›« Ê—… ·ﬂÏ  ﬂÊ‰ „—Ã⁄ ·ﬂ„ ›Ì„« »⁄œ
InvoiceHash = res.InvoiceHash
SingedXML = res.SingedXML
EncodedInvoice = res.EncodedInvoice
 UUID = res.UUID
QRCode = res.QRCode
PIH = res.PIH
SingedXMLFileName = res.SingedXMLFileName

 

' after complete xml send it to zatca api
Dim apireqlogic As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
     Dim tokenresponse As Object
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
     Dim invrequestbody As Object
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
      Dim invoicereportingmodel As Object
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
     invrequestbody.invoice = res.EncodedInvoice
      invrequestbody.InvoiceHash = res.InvoiceHash
      invrequestbody.UUID = res.UUID
      
    
     
     If SendingMode = 0 Then 'dev
     apireqlogic.SetApiRequestLogic 0
         Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
         
    If tokenresponse.errorMessage = "" Then
        Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
 
                                    If invoicereportingmodel.IsSuccess = True Then
                      If invoicereportingmodel.StatusCode = 202 Then
                    '   MsgBox invoicereportingmodel.WarningMessage
                        
                                     WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                      '     str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                                      '       Cn2.Execute str
                       
                       
                      End If
                                 ' reporting status must be REPORTED
                                 ' MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                                  
                                 ErrorMessageS = invoicereportingmodel.ReportingStatus
                                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                                    zatcaStatus = 1
                                                    
                                                    str = " update " & mTableName & " set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                             cn2.Execute str
                                             updatedatatable val(docType), mTableName, mFieldIDName
                   
                                  Else
                                'MsgBox invoicereportingmodel.ErrorMessage
                                ErrorMessageS = invoicereportingmodel.errorMessage
                             savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                           '     str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                           '     Cn2.Execute str
                                End If
  Else
                    'MsgBox tokenresponse.ErrorMessage
                    ErrorMessageS = tokenresponse.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
 'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
 'Cn2.Execute str
 
 
 
                     
                        
                        
End If



     ElseIf SendingMode = 1 Then 'simu
      apireqlogic.SetApiRequestLogic 1
      Set invoicereportingmodel = apireqlogic.CallReportingAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                     If invoicereportingmodel.IsSuccess = True Then
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                             WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), mTableName, mFieldIDName
                        
                           '    str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           '                  Cn2.Execute str
                      End If
                     ' reporting status must be REPORTED
                  '    MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                      ErrorMessageS = invoicereportingmodel.ReportingStatus
                      savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                        zatcaStatus = 1
                              
                                        str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                  cn2.Execute str
 
                      updatedatatable val(docType), mTableName, mFieldIDName
                      Else
                  '  MsgBox invoicereportingmodel.ErrorMessage
                    ErrorMessageS = invoicereportingmodel.errorMessage
                      savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
             '       str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
             '        Cn2.Execute str
                     
                      
                    End If
        ElseIf SendingMode = 2 Then 'prood
        apireqlogic.SetApiRequestLogic 2
        Set invoicereportingmodel = apireqlogic.CallReportingAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                     If invoicereportingmodel.IsSuccess = True Then
                     ' reporting status must be REPORTED
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                                              WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                           '    str = " update Transactions  set  WarningMessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           '                  Cn2.Execute str
                      End If
                  '    MsgBox invoicereportingmodel.ReportingStatus 'REPORTED
                      ErrorMessageS = invoicereportingmodel.ReportingStatus
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                         
                                                zatcaStatus = 1
                                           
                                        str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID


                                  cn2.Execute str
 
 updatedatatable val(docType), mTableName, mFieldIDName
                      Else
                          ErrorMessageS = errorMessage.ReportingStatus
               '     MsgBox invoicereportingmodel.ErrorMessage
                           'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                              savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                     'Cn2.Execute str
                    End If
        End If

   ' End If
  
   
Else
'MsgBox res.ErrorMessage
ErrorMessageS = res.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
       'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
       '              Cn2.Execute str
End If

Else





If res.IsValid = True Then
InvoiceHash = res.InvoiceHash
SingedXML = res.SingedXML
EncodedInvoice = res.EncodedInvoice
 UUID = res.UUID
QRCode = res.QRCode
PIH = res.PIH
SingedXMLFileName = res.SingedXMLFileName

 

' after complete xml send it to zatca api
'Dim apireqlogic As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
     'Dim tokenresponse As Object
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
     'Dim invrequestbody As Object
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
      'Dim invoicereportingmodel As Object
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
     Dim invoiceclearingmodel As Object
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")
     
     invrequestbody.invoice = res.EncodedInvoice
      invrequestbody.InvoiceHash = res.InvoiceHash
      invrequestbody.UUID = res.UUID
      
    ' Dim SendingMode As Integer
     
     If SendingMode = 0 Then 'dev
         Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
    If tokenresponse.errorMessage = "" Then
   
         Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                    If invoicereportingmodel.IsSuccess = True Then
                     
                      If invoicereportingmodel.StatusCode = 202 Then
                       'MsgBox invoicereportingmodel.WarningMessage
                       
                                                                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                        
                        
                             '  str = " update Transactions  set  warrningmessage='" & invoicereportingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                             '                Cn2.Execute str
                      End If
                      
                    
                     ' reporting status must be CLEARED
                   '   MsgBox invoicereportingmodel.ClearanceStatus 'CLEARED
                      ErrorMessageS = invoicereportingmodel.ClearanceStatus
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                                       zatcaStatus = 1
                   
                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                        

                   cn2.Execute str 'WAELSALIM
                   updatedatatable val(docType), mTableName, mFieldIDName
         
 
                      Else
                  '  MsgBox invoicereportingmodel.ErrorMessage
                    ErrorMessageS = invoicereportingmodel.errorMessage
                       savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                           'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                     'Cn2.Execute str
                     
                    End If
  Else
'MsgBox tokenresponse.ErrorMessage
ErrorMessageS = tokenresponse.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
        'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
        '             Cn2.Execute str
End If
     ElseIf SendingMode = 1 Then 'simu
      
      apireqlogic.SetApiRequestLogic 1
     Call apireqlogic.SetToSaveClearedXMLFile(App.path, True)
      
      Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
                If invoiceclearingmodel.IsSuccess = True Then
                     
                      If invoiceclearingmodel.StatusCode = 202 Then
                '       MsgBox invoiceclearingmodel.WarningMessage
                      '  str = " update Transactions  set  warrningmessage='" & invoiceclearingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                           'MsgBox invoiceclearingmodel.WarningMessage
                            warrningmessage = invoiceclearingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(warrningmessage), , mTableName, mFieldIDName
                                         '    Cn2.Execute str
                      End If
                      
              '  MsgBox invoiceclearingmodel.ClearanceStatus 'Cleared
             ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, CStr(WarningMessage), CStr(ErrorMessageS), mTableName, mFieldIDName
                   'MsgBox invoiceclearingmodel.QRCode
                   QRCode = invoiceclearingmodel.QRCode
                   'MsgBox invoiceclearingmodel.ClearedInvoice
                                   zatcaStatus = 1
                   
                  
                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                 

               cn2.Execute str
 updatedatatable val(docType), mTableName, mFieldIDName
                  Else
                'MsgBox invoiceclearingmodel.ErrorMessage
                ErrorMessageS = invoiceclearingmodel.errorMessage
                         savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                       'str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                     'Cn2.Execute str
                End If
        ElseIf SendingMode = 2 Then 'prood
     apireqlogic.SetToSaveClearedXMLFile App.path, True
       apireqlogic.SetApiRequestLogic 2
        Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
     If invoiceclearingmodel.IsSuccess = True Then
                     
                      If invoiceclearingmodel.StatusCode = 202 Then
                    '   MsgBox invoiceclearingmodel.WarningMessage
                    WarningMessage = invoiceclearingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), mTableName, mFieldIDName
                        
                          '  str = " update Transactions  set  warrningmessage='" & invoiceclearingmodel.WarningMessage & "'  where Transaction_ID =" & Transaction_ID
                          '                   Cn2.Execute str
                      End If
                      
'MsgBox invoiceclearingmodel.ClearanceStatus 'Cleared
ErrorMessageS = invoiceclearingmodel.ClearanceStatus
savewarrninganderror Transaction_ID, , CStr(ErrorMessageS)
'   MsgBox invoiceclearingmodel.QRCode
   QRCode = invoiceclearingmodel.QRCode
   
  ' MsgBox invoiceclearingmodel.ClearedInvoice
                                 zatcaStatus = 1
                   'str = " update Transactions  set zatcaStatus  =1    where Transaction_ID =" & Transaction_ID

                                            str = " update " & mTableName & "  set zatcaStatus  =1    where " & mFieldIDName & "  =" & Transaction_ID
                                        

               cn2.Execute str
               updatedatatable val(docType), mTableName, mFieldIDName
  Else
'MsgBox invoiceclearingmodel.ErrorMessage
ErrorMessageS = invoiceclearingmodel.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
         '      str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
                 '    Cn2.Execute str
End If
        End If

   ' End If
  
   
Else
'MsgBox res.ErrorMessage
ErrorMessageS = res.errorMessage
   savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
     '          str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
     '                Cn2.Execute str
End If



End If




'
'Set res = ubl.GenerateInvoiceXML(inv, App.path)
'If res.IsValid = True Then
''MsgBox res.InvoiceHash
''MsgBox res.SingedXML
''MsgBox res.EncodedInvoice
''MsgBox res.UUID
''MsgBox res.QRCode
''MsgBox res.PIH
''MsgBox res.SingedXMLFileName
'
'InvoiceHash = res.InvoiceHash
'SingedXML = res.SingedXML
'EncodedInvoice = res.EncodedInvoice
'UUID = res.UUID
'QRCode = res.QRCode
'PIH = res.PIH
'SingedXMLFileName = res.SingedXMLFileName
'
'updatedatatable
'' after complete xml send it to zatca api
'Dim apireqlogic As ZatcaUBLXMLGenerator.ApiRequestLogic
'    Set apireqlogic = New ZatcaUBLXMLGenerator.ApiRequestLogic
'   Dim tokenresponse As ZatcaUBLXMLGenerator.ComplianceCsrResponse
'    Set tokenresponse = New ZatcaUBLXMLGenerator.ComplianceCsrResponse
'
'     Dim invrequestbody As ZatcaUBLXMLGenerator.InvoiceReportingRequest
'    Set invrequestbody = New ZatcaUBLXMLGenerator.InvoiceReportingRequest
'    Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
'    If tokenresponse.ErrorMessage = "" Then
'
'      invrequestbody.invoice = res.EncodedInvoice
'      invrequestbody.InvoiceHash = res.InvoiceHash
'      invrequestbody.UUID = res.UUID
'         Dim invoicereportingmodel As ZatcaUBLXMLGenerator.InvoiceReportingResponse
'         Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
' If invoicereportingmodel.ErrorMessage = "" Then
'ErrorMessageS = invoicereportingmodel.ReportingStatus
' ' reporting status must be CLEARED
' ' MsgBox invoicereportingmodel.ClearanceStatus 'CLEARED
'                  If invoicereportingmodel.ClearanceStatus = "CLEARED" Or invoicereportingmodel.ReportingStatus = "REPORTED" Then
'
'                   zatcaStatus = 1
'                   str = " update Transactions  set zatcaStatus  =1    where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'
'                If invoicereportingmodel.ClearanceStatus = "CLEARED" Then
'                ErrorMessageS = "CLEARED"
'                    str = " update Transactions  ErrorMessageS=  'CLEARED'   where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'                End If
'
'                       If invoicereportingmodel.ReportingStatus = "REPORTED" Then
'                ErrorMessageS = "REPORTED"
'                    str = " update Transactions  ErrorMessageS=  'REPORTED'   where Transaction_ID =" & Transaction_ID
'                Cn2.Execute str
'
'                End If
'                  End If
'
'  Else
'
'MsgBox invoicereportingmodel.ErrorMessage
'ErrorMessageS = invoicereportingmodel.ErrorMessage
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'
'End If
'  Else
''MsgBox tokenresponse.ErrorMessage
'ErrorMessageS = "ıError in send " & tokenresponse.ErrorMessage & CHR(13) & "INV#" & ID
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'End If
'   ' End If
'
'
'Else
''MsgBox res.ErrorMessage
'ErrorMessageS = "ıError Before send in  " & res.ErrorMessage & CHR(13) & "INV#" & ID
'   str = " update Transactions  set  ErrorMessageS='" & ErrorMessageS & "'  where Transaction_ID =" & Transaction_ID
'Cn2.Execute str
'End If


End Function






Private Sub ConectionFirst()
    On Error GoTo ConnErr
Dim SysSQLServerUserpassword2 As String
SysSQLServerUserpassword2 = "Admin.com"

    If Not POSConnection Is Nothing Then
        If POSConnection.State = adStateOpen Then
            POSConnection.Close
        End If
        Set POSConnection = Nothing
    End If

    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword2 & _
                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                ";Initial Catalog=" & SystemOptions.DbNameW & _
                                ";Data Source=" & SystemOptions.ServerNameW
        ElseIf SysSQLServerType = 2 Then
            If SysSQLServerTypeTechnical = "0" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                                    "Initial Catalog=" & SystemOptions.DbNameW & _
                                    ";Data Source=" & SystemOptions.ServerNameW
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword2 & _
                                    ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                    ";Initial Catalog=" & SystemOptions.DbNameW & _
                                    ";Data Source=" & SystemOptions.ServerNameW
            End If
        End If
        .Open
    End With
    Exit Sub

ConnErr:
    MsgBox "›‘· «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« !" & vbCrLf & Err.Description, vbCritical, "Œÿ√ ›Ì «·« ’«·"
End Sub






' =====================================================
' GenerateInvoice_ZatcaV3 - «·‰”Œ… «·„’ÕÕ… «·‰Â«∆Ì…
' =====================================================
' «·≈’·«Õ«  «·„÷„‰…:
' 1. BR-KSA-F-07/F-13:  ’ÕÌÕ schemeID ··⁄„Ì· «·‰ﬁœÌ (OTH Ê·Ì” TIN)
' 2. BR-CL-14:  ‰ŸÌ› ﬂÊœ «·œÊ·…
' 3. BR-KSA-F-10:  ’ÕÌÕ schemeID ·—ﬁ„ 700 (MOM)
'
' „·«ÕŸ…: BR-KSA-CL-04 (TaxExemptionReasonCode ›Ì ClassifiedTaxCategory)
' Â–«  ‰»ÌÂ Ê·Ì” Œÿ√ - «·„ﬂ »… ·«  œ⁄„ Â–Â «·Œ«’Ì… Õ«·Ì«
' «·›Ê« Ì—  ıﬁ»· „⁄ «· ‰»ÌÂ
' =====================================================

Public Function GenerateInvoice_ZatcaV3( _
    Optional ByVal docType As Integer = 0, _
    Optional ByVal mTableName As String = "Transactions", _
    Optional ByVal mFieldIDName As String = "Transaction_ID", _
    Optional ByVal ActivityTypeIdInv As Integer = 0, _
    Optional ByVal branch_idInv As Integer = 0 _
) As Boolean

On Error Resume Next

    If SystemOptions.PriceWithVAT Then
        Call generateInvoice(docType, mTableName, mFieldIDName, ActivityTypeIdInv, branch_idInv)
        GenerateInvoice_ZatcaV3 = True
        Exit Function
    End If

Dim ubl As Object, inv As Object, res As Object
Set ubl = CreateObject("ZatcaUBLXMLGenerator.UBLXML")
Set inv = CreateObject("ZatcaUBLXMLGenerator.Invoice")
Set res = CreateObject("ZatcaUBLXMLGenerator.Result")

If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    Set cn2 = POSConnection
Else
    Set cn2 = Cn
End If

If SystemOptions.ApplyEinvoiceWithActive Or SystemOptions.ApplyEinvoiceWithBranch Then
    fillmycompanydata ActivityTypeIdInv, branch_idInv
Else
    fillmycompanydata
End If

' ---------------------- BillingReference ----------------------
If (val(InvoiceTypeCodeID) = 381 Or val(InvoiceTypeCodeID) = 383) Then
    If Len(Trim$(InvoiceDocumentReferenceID & "")) > 0 And IsNumeric(InvoiceDocumentReferenceID) Then
        Dim InvoiceDocumentReference As Object
        Set InvoiceDocumentReference = CreateObject("ZatcaUBLXMLGenerator.InvoiceDocumentReference")
        InvoiceDocumentReference.ID = CLng(InvoiceDocumentReferenceID)
        inv.BillingReference.invoiceDocumentReferences.Add InvoiceDocumentReference
    End If
End If

inv.InvoiceTypeCode.ID = val(InvoiceTypeCodeID)
inv.InvoiceTypeCode.Name = InvoiceTypeCodename

If DocumentCurrencyCode = "" Or DocumentCurrencyCode = "SR" Then DocumentCurrencyCode = "SAR"
If TaxCurrencyCode = "" Or TaxCurrencyCode = "SR" Then TaxCurrencyCode = "SAR"
inv.DocumentCurrencyCode = DocumentCurrencyCode
inv.TaxCurrencyCode = TaxCurrencyCode

inv.AdditionalDocumentReferencePIH.EmbeddedDocumentBinaryObject = CalllastPIH(uuidCounter)

Dim icvStr As String, icvOnlyDigits As String, i As Long, Ch As String
icvOnlyDigits = ""
icvStr = Trim$(uuidCounter & "")
If icvStr = "" Then icvStr = "1"
For i = 1 To Len(icvStr)
    Ch = mId$(icvStr, i, 1)
    If Ch >= "0" And Ch <= "9" Then icvOnlyDigits = icvOnlyDigits & Ch
Next i
If icvOnlyDigits = "" Then icvOnlyDigits = "1"
inv.AdditionalDocumentReferenceICV.UUID = icvOnlyDigits

inv.ID = ID

If IsNull(IssueDate) Or Trim$(IssueDate & "") = "" Then IssueDate = Date
inv.IssueDate = Format(IssueDate, "yyyy-mm-dd")

If IsNull(IssueTim) Or Trim$(IssueTim & "") = "" Then IssueTim = Now
inv.IssueTime = Format(IssueTim, "HH:mm:ss")

If Len(Trim$(ActualDeliveryDate & "")) = 0 Then ActualDeliveryDate = IssueDate
If Len(Trim$(LatestDeliveryDate & "")) = 0 Then LatestDeliveryDate = IssueDate
inv.Delivery.ActualDeliveryDate = Format(ActualDeliveryDate, "yyyy-mm-dd")
If val(InvoiceTypeCodeID) = 388 Then
    inv.Delivery.LatestDeliveryDate = Format(LatestDeliveryDate, "yyyy-mm-dd")
End If

Dim paymentmean As Object
Set paymentmean = CreateObject("ZatcaUBLXMLGenerator.PaymentMeans")
paymentmean.PaymentMeansCode = Trim$(PaymentMeansCode & "")
If val(InvoiceTypeCodeID) = 383 Or val(InvoiceTypeCodeID) = 381 Then
    paymentmean.InstructionNote = "Returned items"
End If
If Len(Trim$(PayeeFinancialAccount & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.ID = Trim$(PayeeFinancialAccount & "")
End If
If Len(Trim$(paymentnote & "")) > 0 Then
    paymentmean.PayeeFinancialAccount.paymentnote = Trim$(paymentnote & "")
End If
inv.PaymentMeans.Add paymentmean

' ---------------------- Supplier ----------------------
inv.SupplierParty.PartyIdentification.ID = Trim$(cIdentificationID & "")
If Len(Trim$(cschemeID & "")) = 0 Then cschemeID = "CRN"
inv.SupplierParty.PartyIdentification.schemeID = cschemeID

inv.SupplierParty.PostalAddress.StreetName = IIf(Len(Trim$(cStreetName & "")) = 0, "N/A", Trim$(cStreetName & ""))
inv.SupplierParty.PostalAddress.AdditionalStreetName = Trim$(cAdditionalStreetName & "")
inv.SupplierParty.PostalAddress.BuildingNumber = IIf(Len(Trim$(cBuildingNumber & "")) = 0, "N/A", Trim$(cBuildingNumber & ""))
inv.SupplierParty.PostalAddress.PlotIdentification = Trim$(cPlotIdentification & "")
inv.SupplierParty.PostalAddress.CityName = IIf(Len(Trim$(cCityName & "")) = 0, "N/A", Trim$(cCityName & ""))
inv.SupplierParty.PostalAddress.PostalZone = IIf(Len(Trim$(cPostalZone & "")) = 0, "00000", Trim$(cPostalZone & ""))
inv.SupplierParty.PostalAddress.CountrySubentity = Trim$(cCountrySubentity & "")
inv.SupplierParty.PostalAddress.CitySubdivisionName = Trim$(cCitySubdivisionName & "")

Dim supplierCountry As String
supplierCountry = CleanCountryCode(cIdentificationCode)
inv.SupplierParty.PostalAddress.Country.IdentificationCode = supplierCountry

inv.SupplierParty.PartyLegalEntity.RegistrationName = Trim$(Companynameforvat & "")
inv.SupplierParty.PartyTaxScheme.CompanyID = Trim$(cCompanyID & "")

' =========================================================
' ============   »Ì«‰«  «·⁄„Ì· (Customer)   ==============
' =========================================================
Dim isStandard As Boolean
Dim isSimplified As Boolean

isStandard = (mId$(InvoiceTypeCodename & "", 1, 2) = "01")
isSimplified = (mId$(InvoiceTypeCodename & "", 1, 2) = "02") _
               And (LCase$(mTableName) = "tbleinvoice" _
                 Or LCase$(mTableName) = "propertyduebatchdetail" _
                 Or LCase$(mTableName) = "debitandcreditnotification")

If isStandard Or isSimplified Then
    Dim CUSTID As String
    Dim custSchemeID As String
    
    ' ===  ÕœÌœ —ﬁ„ «·⁄„Ì· ===
    If Len(Trim$(Id700 & "")) > 0 Then
        CUSTID = Trim$(Id700 & "")
        custSchemeID = "MOM"  ' —ﬁ„ 700 œ«∆„« MOM
    Else
        CUSTID = Trim$(Identificationid & "")
        custSchemeID = Trim$(schemeID & "")
    End If
    
    ' =========================================================
    ' ===   ≈’·«Õ BR-KSA-F-07 / BR-KSA-F-13   ================
    ' ===   «·—ﬁ„ «·«› —«÷Ì ÌÃ» √‰ ÌﬂÊ‰ OTH Ê·Ì” TIN!   =====
    ' =========================================================
    If isStandard Then
        If CUSTID = "300000000000003" Or CUSTID = "" Or CUSTID = "0" Or UCase$(CUSTID) = "N/A" Then
            ' Õ«Ê· «” Œœ«„ «·—ﬁ„ «·÷—Ì»Ì √Ê·«
            If Len(Trim$(CompanyID & "")) = 15 And IsNumeric(CompanyID) Then
                ' «·⁄„Ì· ·Â —ﬁ„ ÷—Ì»Ì ’«·Õ
                CUSTID = Trim$(CompanyID)
                custSchemeID = "TIN"
            Else
                ' ===== ⁄„Ì· ‰ﬁœÌ =====
                ' «” Œœ„ «·—ﬁ„ «·«› —«÷Ì „⁄ OTH (Ê·Ì” TIN!)
                CUSTID = "300000000000003"
                custSchemeID = "OTH"  ' <<<< „Â„ Ãœ«!
            End If
        End If
    Else
        ' ›« Ê—… „»”ÿ…
        If Len(CUSTID) = 0 Or UCase$(CUSTID) = "N/A" Or CUSTID = "0" Then
            CUSTID = "300000000000003"
            custSchemeID = "OTH"
        End If
    End If
    
    ' ===  ’ÕÌÕ schemeID  ·ﬁ«∆Ì« (BR-KSA-F-10) ===
    custSchemeID = GetCorrectSchemeID(CUSTID, custSchemeID)
    
    inv.CustomerParty.PartyIdentification.ID = CUSTID
    inv.CustomerParty.PartyIdentification.schemeID = custSchemeID

    ' === ⁄‰Ê«‰ «·⁄„Ì· ===
    Dim cStreet$, cAddStreet$, cBldg$, cCity$, cPost$, cSubent$, cSubCity$, cCountry$
    cStreet = IIf(Len(Trim$(StreetName & "")) = 0, "N/A", Trim$(StreetName & ""))
    cAddStreet = Trim$(AdditionalStreetName & "")
    cBldg = IIf(Len(Trim$(BuildingNumber & "")) = 0, "N/A", Trim$(BuildingNumber & ""))
    cCity = IIf(Len(Trim$(CityName & "")) = 0, "N/A", Trim$(CityName & ""))
    cPost = IIf(Len(Trim$(PostalZone & "")) = 0, "00000", Trim$(PostalZone & ""))
    cSubent = Trim$(CountrySubentity & "")
    cSubCity = Trim$(CitySubdivisionName & "")
    
    ' === ≈’·«Õ ﬂÊœ œÊ·… «·⁄„Ì· (BR-CL-14) ===
    cCountry = CleanCountryCode(IdentificationCode)

    inv.CustomerParty.PostalAddress.StreetName = cStreet
    inv.CustomerParty.PostalAddress.AdditionalStreetName = cAddStreet
    inv.CustomerParty.PostalAddress.BuildingNumber = cBldg
    inv.CustomerParty.PostalAddress.PlotIdentification = Trim$(PlotIdentification & "")
    inv.CustomerParty.PostalAddress.CityName = cCity
    inv.CustomerParty.PostalAddress.PostalZone = cPost
    inv.CustomerParty.PostalAddress.CountrySubentity = cSubent
    inv.CustomerParty.PostalAddress.CitySubdivisionName = cSubCity
    inv.CustomerParty.PostalAddress.Country.IdentificationCode = cCountry

    inv.CustomerParty.PartyLegalEntity.RegistrationName = Trim$(RegistrationName & "")
    
    If Len(Trim$(Id700 & "")) > 0 Then
        inv.CustomerParty.PartyTaxScheme.CompanyID = ""
    Else
        inv.CustomerParty.PartyTaxScheme.CompanyID = Trim$(CompanyID & "")
    End If
End If

' =========================================================
' ============   «·»‰Êœ   ================================
' =========================================================
Dim StrSQL As String
StrSQL = ""
StrSQL = StrSQL & "SELECT "
StrSQL = StrSQL & "  I.ItemName, I.ItemNamee, ISNULL(I.ItemWithOutVAT,0) AS ItemWithOutVAT, "
StrSQL = StrSQL & "  T.Trans_DiscountType, T.Trans_Discount, "
StrSQL = StrSQL & "  TD.ShowQty, "
StrSQL = StrSQL & "  TD.ItemDiscountType, "
StrSQL = StrSQL & "  ISNULL(TD.ItemDiscount,0) AS ItemDiscount, "
StrSQL = StrSQL & "  ISNULL(TD.TotalDiscountPerLine,0) AS TotalDiscountPerLine, "
StrSQL = StrSQL & "  VatRate = CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END, "

If SystemOptions.PriceWithVAT Then
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND((TD.showPrice * TD.ShowQty) / (1 + (CASE WHEN ISNULL(I.ItemWithOutVAT,0)=0 THEN ISNULL(TD.Vatyo,0) ELSE 0 END)/100.0), 6) "
Else
    StrSQL = StrSQL & "  BasePrice_Ex = ROUND(TD.showPrice, 6), "
    StrSQL = StrSQL & "  LineEx_NoDisc = ROUND(TD.showPrice * TD.ShowQty, 6) "
End If

StrSQL = StrSQL & "FROM dbo.Transactions AS T "
StrSQL = StrSQL & "INNER JOIN dbo.Transaction_Details AS TD ON T.Transaction_ID = TD.Transaction_ID "
StrSQL = StrSQL & "INNER JOIN dbo.TblItems AS I ON TD.Item_ID = I.ItemID "
StrSQL = StrSQL & "WHERE T.Transaction_ID = " & Transaction_ID

Dim rs As New ADODB.Recordset
rs.Open StrSQL, cn2, adOpenStatic, adLockReadOnly, adCmdText

Dim SumExNoDisc As Double
Dim SumTaxableExNoDisc As Double
Dim SumExemptExNoDisc As Double
Dim SumLineDiscEx As Double
Dim mTrans_DiscountType As Integer
Dim mTrans_Discount As Double

SumExNoDisc = 0#
SumTaxableExNoDisc = 0#
SumExemptExNoDisc = 0#
SumLineDiscEx = 0#

If Not rs.EOF Then
    mTrans_DiscountType = val(rs!Trans_DiscountType & "")
    mTrans_Discount = val(rs!Trans_Discount & "")
End If

Do Until rs.EOF
    Dim lineEx As Double
    Dim Rate As Double
    Dim Qty As Double
    Dim basePrice As Double
    Dim discType As Long
    Dim discPercent As Double
    Dim discValue As Double
    Dim discRaw As Double
    Dim lineDiscEx As Double
    Dim lineBaseForPercent As Double
    Dim isItemExempt As Boolean
    Dim ratePct As Double
    
    lineEx = val(rs!LineEx_NoDisc & "")
    Rate = val(rs!vatRate & "")
    Qty = val(rs!ShowQty & "")
    basePrice = val(rs!BasePrice_Ex & "")
    isItemExempt = (val(rs!ItemWithOutVAT & "") <> 0)

    ratePct = IIf(Rate > 1, Rate, Rate * 100)

    discType = val(rs!ItemDiscountType & "")
    discPercent = val(rs!ItemDiscount & "")
    discValue = val(rs!ItemDiscount & "")
    discRaw = 0

    Select Case discType
        Case 3
            If SystemOptions.PriceWithVAT And Rate > 0 Then
                lineBaseForPercent = lineEx * (1 + Rate / 100#)
            Else
                lineBaseForPercent = lineEx
            End If
            discRaw = Round(lineBaseForPercent * discPercent / 100#, 6) + val(rs!TotalDiscountPerLine & "")
        Case 2
            discRaw = discValue + val(rs!TotalDiscountPerLine & "")
        Case Else
            discRaw = val(rs!TotalDiscountPerLine & "")
    End Select

    If SystemOptions.PriceWithVAT And Rate > 0 Then
        lineDiscEx = Round(discRaw / (1 + (Rate / 100#)), 6)
    Else
        lineDiscEx = Round(discRaw, 6)
    End If

    SumExNoDisc = SumExNoDisc + lineEx

    If ratePct > 0 And Not isItemExempt Then
        SumTaxableExNoDisc = SumTaxableExNoDisc + lineEx
    Else
        SumExemptExNoDisc = SumExemptExNoDisc + lineEx
    End If

    SumLineDiscEx = SumLineDiscEx + lineDiscEx

    Dim invline As Object
    Set invline = CreateObject("ZatcaUBLXMLGenerator.InvoiceLine")
    
    If docType = 0 Then
        invline.Price.EncludingVat = SystemOptions.PriceWithVAT
    Else
        invline.Price.EncludingVat = False
    End If
    
    invline.Price.PriceAmount = Format(basePrice, "0.######")
    invline.InvoiceQuantity = Format(Qty, "0.######")
    invline.Item.Name = rs!ItemName & ""

    ' =========================================================
    ' ===   «· ’‰Ì› «·÷—Ì»Ì   ================================
    ' =========================================================
    Dim taxCatID As String
    Dim exemptCode As String
    Dim exemptReason As String
    
    taxCatID = "S"
    exemptCode = ""
    exemptReason = ""
    
    ' === «·⁄Ì‰«  «·„Ã«‰Ì… ===
    If Round(basePrice, 2) = 0 Then
        taxCatID = "Z"
        exemptCode = "VATEX-SA-00"
        exemptReason = "Free sample"
        
    ElseIf ratePct > 0 And Not isItemExempt Then
        taxCatID = "S"
        
    Else
        If docType = 6 And isItemExempt Then
            taxCatID = "E"
            exemptCode = "VATEX-SA-30"
            exemptReason = "Real estate residential lease/sale"
            
        ElseIf InvoiceTypeCodename = "0100100" Then
            taxCatID = "Z"
            exemptCode = "VATEX-SA-32"
            exemptReason = "Export of goods"
            
        ElseIf InvoiceTypeCodename = "0100000" And docType = 6 Then
            taxCatID = "O"
            exemptCode = "VATEX-SA-OOS"
            exemptReason = "Out of scope of VAT"
            
        ElseIf LCase$(mTableName) = "debitandcreditnotification" And ratePct = 0 Then
            taxCatID = "O"
            exemptCode = "VATEX-SA-OOS"
            exemptReason = "Out of scope of VAT"
            
        Else
            taxCatID = "E"
            exemptCode = "VATEX-SA-29"
            exemptReason = "Financial services specified in Article 29"
        End If
    End If
    
    '  ⁄ÌÌ‰ «· ’‰Ì›
    invline.Item.classifiedTaxCategory.ID = taxCatID
    invline.Item.classifiedTaxCategory.Percent = IIf(taxCatID = "S", ratePct, 0)
    
    ' „Õ«Ê·… ≈÷«›… TaxExemptionReasonCode (ﬁœ ·«  œ⁄„Â« «·„ﬂ »…)
    On Error Resume Next
    If exemptCode <> "" Then
        invline.Item.classifiedTaxCategory.TaxExemptionReasonCode = exemptCode
        invline.Item.classifiedTaxCategory.TaxExemptionReason = exemptReason
    End If
    Err.Clear
    On Error Resume Next
    
    invline.TaxTotal.TaxSubtotal.taxCategory.ID = taxCatID
    invline.TaxTotal.TaxSubtotal.taxCategory.Percent = IIf(taxCatID = "S", ratePct, 0)
    
    If exemptCode <> "" Then
        On Error Resume Next
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReasonCode = exemptCode
        invline.TaxTotal.TaxSubtotal.taxCategory.TaxExemptionReason = exemptReason
        Err.Clear
        On Error Resume Next
    End If

    inv.InvoiceLines.Add invline

    rs.MoveNext
Loop

rs.Close

' =========================================================
' ==============  «·Œ’Ê„«  Ê«·≈Ã„«·Ì«   ==================
' =========================================================
Dim HeaderDiscEx As Double
Dim TransDiscountRaw As Double
Dim TransDiscountEx As Double

HeaderDiscEx = Round(SumLineDiscEx, 2)
TransDiscountRaw = val(Trans_Discount & "")

If TransDiscountRaw <> 0 Then
    If val(mTrans_DiscountType) = 1 Then
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            Dim TotalWithVAT As Double
            TotalWithVAT = SumExNoDisc * (1 + 0.15)
            TransDiscountRaw = Round(TotalWithVAT * TransDiscountRaw / 100#, 6)
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            TransDiscountRaw = Round(SumExNoDisc * TransDiscountRaw / 100#, 6)
            TransDiscountEx = TransDiscountRaw
        End If
    Else
        If SystemOptions.PriceWithVAT And SumTaxableExNoDisc > 0 Then
            TransDiscountEx = Round(TransDiscountRaw / 1.15, 6)
        Else
            TransDiscountEx = TransDiscountRaw
        End If
    End If
    HeaderDiscEx = HeaderDiscEx + TransDiscountEx
End If

If HeaderDiscEx < 0 Then HeaderDiscEx = 0

Dim BaseTotal As Double: BaseTotal = SumExNoDisc
Dim Allow_S As Double, Allow_E As Double

If BaseTotal > 0 Then
    Allow_S = Round(HeaderDiscEx * (SumTaxableExNoDisc / BaseTotal), 2)
    Allow_E = Round(HeaderDiscEx - Allow_S, 2)
Else
    Allow_S = 0
    Allow_E = HeaderDiscEx
End If

Dim HeaderDiscVat As Double
HeaderDiscVat = Round(Allow_S * 0.15, 2)

If HeaderDiscEx > 0# Then
    If Allow_S > 0# Then
        Dim allowanceS As Object
        Set allowanceS = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceS.ChargeIndicator = False
        allowanceS.Amount = Format(Allow_S, "0.##")
        allowanceS.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceS.taxCategory.ID = "S"
        allowanceS.taxCategory.Percent = 15#
        inv.allowanceCharges.Add allowanceS
    End If

    If Allow_E > 0# Then
        Dim allowanceE As Object
        Set allowanceE = CreateObject("ZatcaUBLXMLGenerator.AllowanceCharge")
        allowanceE.ChargeIndicator = False
        allowanceE.Amount = Format(Allow_E, "0.##")
        allowanceE.AllowanceChargeReason = "Invoice discount (unified)"
        allowanceE.taxCategory.ID = "E"
        allowanceE.taxCategory.Percent = 0#
        inv.allowanceCharges.Add allowanceE
    End If
End If

SumExNoDisc = Round(SumExNoDisc, 2)
SumTaxableExNoDisc = Round(SumTaxableExNoDisc, 2)
HeaderDiscEx = Round(HeaderDiscEx, 2)

inv.LegalMonetaryTotal.LineExtensionAmount = Format(SumExNoDisc, "0.##")
inv.LegalMonetaryTotal.AllowanceTotalAmount = Format(HeaderDiscEx + HeaderDiscVat, "0.##")

Dim TaxExclusiveAfterDisc As Double
TaxExclusiveAfterDisc = Round(SumExNoDisc - HeaderDiscEx, 2)
If TaxExclusiveAfterDisc < 0 Then TaxExclusiveAfterDisc = 0

inv.LegalMonetaryTotal.TaxExclusiveAmount = Format(TaxExclusiveAfterDisc, "0.##")

Dim TaxableAfterDisc As Double
Dim TotalVat As Double

TaxableAfterDisc = Round(SumTaxableExNoDisc - Allow_S, 2)
If TaxableAfterDisc < 0 Then TaxableAfterDisc = 0

Dim rawVat As Double
rawVat = TaxableAfterDisc * 0.15
TotalVat = Fix(rawVat * 100) / 100

inv.TaxTotal.TaxAmount = Format(TotalVat, "0.##")

Dim TaxInclusive As Double
TaxInclusive = CDbl(Format(TaxExclusiveAfterDisc, "0.00")) + CDbl(Format(TotalVat, "0.00"))
TaxInclusive = Round(TaxInclusive, 2)

inv.LegalMonetaryTotal.TaxInclusiveAmount = Format(TaxInclusive, "0.##")
inv.LegalMonetaryTotal.ChargeTotalAmount = "0.00"
inv.LegalMonetaryTotal.PayableRoundingAmount = "0.00"

Dim Prepaid As Double
If val(InvoiceTypeCodeID) = 388 Then
    Prepaid = val(PrepaidAmount & "")
Else
    Prepaid = 0
End If
inv.LegalMonetaryTotal.PrepaidAmount = Format(Round(Prepaid, 2), "0.##")

Dim PayableAmountCalc As Double
PayableAmountCalc = Round(TaxInclusive - Prepaid, 2)
If PayableAmountCalc < 0 Then PayableAmountCalc = 0
inv.LegalMonetaryTotal.PayableAmount = Format(PayableAmountCalc, "0.##")

' =========================================================
' ==================   «·≈—”«·   ==========================
' =========================================================
Dim SendingMode As Integer: SendingMode = SystemOptions.SendingMode
Dim SecretKey As String:    SecretKey = SystemOptions.SecretKey

If SendingMode <> 0 Then
    inv.CSIDInfo.CertPem = SystemOptions.PublickeycertPem
    inv.CSIDInfo.Privatekey = SystemOptions.Privatekey
End If

Set res = ubl.GenerateInvoiceXML(inv, App.path)

If res.IsValid = True Then
    InvoiceHash = res.InvoiceHash
    SingedXML = res.SingedXML
    EncodedInvoice = res.EncodedInvoice
    UUID = res.UUID
    QRCode = res.QRCode
    PIH = res.PIH
    SingedXMLFileName = res.SingedXMLFileName

    Dim apireqlogic As Object, tokenresponse As Object
    Dim invrequestbody As Object, invoicereportingmodel As Object, invoiceclearingmodel As Object
    Set apireqlogic = CreateObject("ZatcaUBLXMLGenerator.BLL.ApiRequestLogic")
    Set tokenresponse = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.ComplianceCsrResponse")
    Set invrequestbody = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingRequest")
    Set invoicereportingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceReportingResponse")
    Set invoiceclearingmodel = CreateObject("ZatcaUBLXMLGenerator.HelperContracts.InvoiceClearanceResponse")

    invrequestbody.invoice = res.EncodedInvoice
    invrequestbody.InvoiceHash = res.InvoiceHash
    invrequestbody.UUID = res.UUID

    Select Case SendingMode
        Case 0
            apireqlogic.SetApiRequestLogic 0
            Set tokenresponse = apireqlogic.GetComplianceCSIDAPI("12345", "", App.path)
            If tokenresponse.errorMessage = "" Then
                Set invoicereportingmodel = apireqlogic.CallComplianceInvoiceAPI(tokenresponse.BinarySecurityToken, tokenresponse.Secret, invrequestbody)
                If invoicereportingmodel.IsSuccess = True Then
                    If invoicereportingmodel.StatusCode = 202 Then
                        WarningMessage = invoicereportingmodel.WarningMessage
                        savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                    End If
                    ErrorMessageS = invoicereportingmodel.ReportingStatus
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                    zatcaStatus = 1
                    cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                    updatedatatable val(docType), mTableName, mFieldIDName
                Else
                    ErrorMessageS = invoicereportingmodel.errorMessage
                    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                End If
            Else
                ErrorMessageS = tokenresponse.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 1
            apireqlogic.SetApiRequestLogic 1
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If

        Case 2
            apireqlogic.SetApiRequestLogic 2
            apireqlogic.SetToSaveClearedXMLFile App.path, True
            Set invoiceclearingmodel = apireqlogic.CallClearanceAPI(apireqlogic.ToBase64Encode(inv.CSIDInfo.CertPem), SecretKey, invrequestbody)
            If invoiceclearingmodel.IsSuccess = True Then
                If invoiceclearingmodel.StatusCode = 202 Then
                    WarningMessage = invoiceclearingmodel.WarningMessage
                    savewarrninganderror Transaction_ID, CStr(WarningMessage), , mTableName, mFieldIDName
                End If
                ErrorMessageS = invoiceclearingmodel.ClearanceStatus
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
                QRCode = invoiceclearingmodel.QRCode
                zatcaStatus = 1
                cn2.Execute "UPDATE " & mTableName & " SET zatcaStatus=1 WHERE " & mFieldIDName & "=" & Transaction_ID
                updatedatatable val(docType), mTableName, mFieldIDName
            Else
                ErrorMessageS = invoiceclearingmodel.errorMessage
                savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
            End If
    End Select

    GenerateInvoice_ZatcaV3 = True
Else
    ErrorMessageS = res.errorMessage
    savewarrninganderror Transaction_ID, , CStr(ErrorMessageS), mTableName, mFieldIDName
    GenerateInvoice_ZatcaV3 = False
End If

End Function


' =====================================================
' œ«·…  ‰ŸÌ› ﬂÊœ «·œÊ·… (BR-CL-14)
' =====================================================
Private Function CleanCountryCode(ByVal rawCode As String) As String
    Dim cleanCode As String
    Dim j As Integer
    Dim c As String
    
    cleanCode = ""
    rawCode = Trim$(rawCode & "")
    
    For j = 1 To Len(rawCode)
        c = mId$(rawCode, j, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            cleanCode = cleanCode & UCase$(c)
        End If
    Next j
    
    '  ’ÕÌÕ «·√Œÿ«¡ «·‘«∆⁄…
    If cleanCode = "AS" Then cleanCode = "SA"
    
    If Len(cleanCode) = 2 Then
        CleanCountryCode = cleanCode
    Else
        CleanCountryCode = "SA"
    End If
End Function


' =====================================================
' œ«·…  ’ÕÌÕ schemeID (BR-KSA-F-10)
' =====================================================
Private Function GetCorrectSchemeID(ByVal BuyerID As String, ByVal currentSchemeID As String) As String
    BuyerID = Trim$(BuyerID & "")
    currentSchemeID = Trim$(UCase$(currentSchemeID & ""))
    
    ' ===== «·—ﬁ„ «·«› —«÷Ì œ«∆„« OTH =====
    If BuyerID = "300000000000003" Then
        GetCorrectSchemeID = "OTH"
        Exit Function
    End If
    
    If Len(BuyerID) = 0 Then
        GetCorrectSchemeID = "OTH"
        Exit Function
    End If
    
    Dim firstChar As String
    firstChar = left$(BuyerID, 1)
    
    Dim idLength As Integer
    idLength = Len(BuyerID)
    
    ' 15 —ﬁ„ = TIN
    If idLength = 15 And IsNumeric(BuyerID) Then
        GetCorrectSchemeID = "TIN"
        Exit Function
    End If
    
    ' 10 √—ﬁ«„
    If idLength = 10 And IsNumeric(BuyerID) Then
        Select Case firstChar
            Case "1"
                If currentSchemeID = "NAT" Or currentSchemeID = "CRN" Then
                    GetCorrectSchemeID = currentSchemeID
                Else
                    GetCorrectSchemeID = "CRN"
                End If
                
            Case "2"
                If currentSchemeID = "NAT" Or currentSchemeID = "IQA" Then
                    GetCorrectSchemeID = currentSchemeID
                Else
                    GetCorrectSchemeID = "IQA"
                End If
                
            Case "7"
                ' ===== —ﬁ„ 700 = MOM =====
                GetCorrectSchemeID = "MOM"
                
            Case Else
                GetCorrectSchemeID = "CRN"
        End Select
        Exit Function
    End If
    
    If currentSchemeID = "" Then
        GetCorrectSchemeID = "OTH"
    Else
        GetCorrectSchemeID = currentSchemeID
    End If
End Function


