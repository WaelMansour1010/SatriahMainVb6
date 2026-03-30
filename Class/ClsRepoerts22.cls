VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsRepoerts22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 

  
     
 

Public Function ShowSallingTime(StrSQL As String, _
                                Optional DateFrom As Variant = Null, _
                                Optional DateTo As Variant = Null, _
                                Optional BolShowNotes As Boolean = False, _
                                Optional newReport As Variant, _
                                Optional PaymentType As Integer = 2, _
                                Optional reportString As String, _
                                Optional UnitName As String, _
                                Optional ViewCost As String, _
                                Optional reporttitle As String = "", _
                                Optional ReturnPurcahse As Integer = 0, _
                                Optional showvalue As Boolean = False, Optional branchname As String = "", Optional ShowZero As String, Optional Departement As String, Optional GroupName As String, Optional ShowGrouping As Integer = 1, Optional CustomerLocation As String, Optional ByVal mTypeDebitCredit As Integer = 0, Optional mReportName As String = "", Optional salesflag As String, Optional doneflag As Integer = 0)
Dim isGrouping As Integer
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim Reportname As String
    Dim ShowTax As Boolean
    Dim RsNotes As ADODB.Recordset
    Dim MySQL As String
    Dim StrTitle As String
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim i As Integer
    Dim StrFileName As String

    On Error GoTo ErrTrap

     If StrSQL = "" Then
        Exit Function
       End If

    'Œ«’ »÷—Ì»…«·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    ShowTax = False
    
    If ShowTax = True Then
        Reportname = "ReportSalingTime.rpt"
    ElseIf BolShowNotes = False Then
        Reportname = "ReportSalingTimeNotTax.rpt"
    ElseIf BolShowNotes = True Then
        Reportname = "ReportSalingTimeNotTaxDebit.rpt"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«  ÊÃœ »Ì«‰«  ›Ì Â–Â «·› —…"
 Else
 Msg = "No Data in this Period"
 End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If BolShowNotes = True Then
        StrTitle = " ﬁ—Ì— »«·„»Ì⁄«  «·√Ã·… Ê«· Õ’Ì·«  «·ﬁ«∆„… ⁄·ÌÂ«"
        Set RsNotes = New ADODB.Recordset
        MySQL = "Select * From QryPostedTransNotes Where NoteType=4 OR NoteType=9"
        RsNotes.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrTitle = " ﬁ—Ì— »«·„»Ì⁄« "
    Else
    StrTitle = "Sales report "
    End If
        '    If Not (IsNull(DateFrom)) Then
        '        StrTitle = StrTitle & " ›Ï «·› —… „‰ " & DisplayDate(CDate(DateFrom)) & "¯"
        '    End If
        '    If Not (IsNull(DateTo)) Then
        '        StrTitle = StrTitle & " ≈·Ï " & DisplayDate(CDate(DateTo))
        '    End If
        '    StrTitle = StrTitle & "¯"
    End If

    If newReport = 100312 Then
        Reportname = "Inventory\sales100312.rpt"
ElseIf newReport = 1 Then
        Reportname = "Inventory\sales2.rpt"
    ElseIf newReport = 2 Then
        Reportname = "Inventory\salesByBranch.rpt"
    ElseIf newReport = 3 Then
        Reportname = "Inventory\salesByBranchYearly.rpt"
    ElseIf newReport = 4 Then
        Reportname = "Inventory\sales3.rpt"
        StrTitle = " ﬁ—Ì— „»Ì⁄«  ﬂ· «·„Œ«“‰ √›ﬁÌ"
    ElseIf newReport = 5 Then
        Reportname = "Inventory\ItemReport.rpt"
        StrTitle = " ﬁ—Ì— «·«’‰«›"
    ElseIf newReport = 6 Then
        Reportname = "Inventory\Gard4.rpt"
        StrTitle = " ﬁ—Ì—„ «»⁄Â  €Ì— «”⁄«— «·«’‰«›"
    ElseIf newReport = 7 Then
        Reportname = "Inventory\sales4.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    «·«’‰«›"

    ElseIf newReport = 8 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "CommercialTransactions\CommercialTransactions.rpt"
Else
 Reportname = "CommercialTransactions\CommercialTransactionse.rpt"
End If


            If SystemOptions.UserInterface = ArabicInterface Then
        StrTitle = "„—«Ã⁄Â «·„⁄«„·«  «· Ã«—Ì… "
        Else
        StrTitle = "Commercial Transaction Auditing"
        End If

If doneflag = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & " - «·ﬂ·   "
            Else
                    StrTitle = StrTitle & "-Over All"
            End If



ElseIf doneflag = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & " - «·„”·„ "
            Else
                    StrTitle = StrTitle & "  -Delivered"
            End If

Else

            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & "- €Ì— «·”·„ "
            Else
                    StrTitle = "-Not Delivered"
            End If
End If


If Departement <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··«œ«—… :" & Departement
                Else
                          StrTitle = StrTitle & " Departement :" & Departement
                End If

End If



    ElseIf newReport = 845 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "CommercialTransactions\CommercialTransactions1.rpt"
     Else
     Reportname = "CommercialTransactions\CommercialTransactions1e.rpt"
     End If
     StrTitle = "„—«Ã⁄Â «·„⁄«„·«  «· Ã«—Ì… "
        If Departement <> "" Then

                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··«œ«—… :" & Departement
                Else
                          StrTitle = StrTitle & " Departement :" & Departement
                End If

End If

If doneflag = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & "«·ﬂ·   "
            Else
                    StrTitle = StrTitle & "Over All"
            End If



ElseIf doneflag = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & "  «·”·„ "
            Else
                    StrTitle = StrTitle & "  Delivered"
            End If

Else

            If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = StrTitle & "€Ì— «·”·„ "
            Else
                    StrTitle = "Not Delivered"
            End If
End If



    ElseIf newReport = 9 Then
    If UnitName <> "" Then
     StrTitle = " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì… ··„‰œÊ» " & CHR(13) & UnitName
    Else
     StrTitle = " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì…   "
     
    End If
    
        Reportname = "AccountsTransactions\AccountsTransactions.rpt"
 
    ElseIf newReport = 10 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales5.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
    Else
    Reportname = "Inventory\sales5E.rpt"
        StrTitle = "Sales Report Detailed By Unit :" & UnitName
    
    End If
    
    
    ElseIf newReport = 910 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales910.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  Õ·Ì·Ì       " & UnitName
Else

        Reportname = "Inventory\sales910E.rpt"
        StrTitle = "Sales Report Detailed by Items  "

End If

    ElseIf newReport = 1001 Then
        Reportname = "Inventory\sales1001.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ… " & UnitName
    ElseIf newReport = 1002 Then
        Reportname = "Inventory\sales1002.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì   "
    ElseIf newReport = 1003 Then
        Reportname = "Inventory\sales1003.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì   "

    ElseIf newReport = 18000 Then
        Reportname = "Inventory\sales" & newReport & ".rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ… " & UnitName
    ElseIf newReport = 18001 Then
         Reportname = "Inventory\sales" & newReport & ".rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ… " & UnitName
    ElseIf newReport = 18002 Then
         Reportname = "Inventory\sales" & newReport & ".rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ… " & UnitName
        
    ElseIf newReport = 32 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales32.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales32E.rpt"
        StrTitle = "Purcahse Totals By Items For Unit   " & UnitName
 End If
    
    
    ElseIf newReport = 63 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales63.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales63E.rpt"
        StrTitle = "Purcahse Retuen Totals By Items For Unit   " & UnitName
 End If
 
 
 
    ElseIf newReport = 64 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales64.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„»Ì⁄«    ··«’‰«›   Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales64.rpt"
        StrTitle = "Sales Retuen Totals By Items big Unit   " & UnitName
 End If
 
 
   ElseIf newReport = 964 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales964.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„»Ì⁄«    ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ…  «·ﬂ»—Ì      " & UnitName
 Else
 
     Reportname = "Inventory\sales964.rpt"
        StrTitle = "Purcahse Retuen Totals By Items For Unit   " & UnitName
 End If
 
 
 
    ElseIf newReport = 65 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales65.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„»Ì⁄«    ··⁄„·«¡    »«·ÊÕœ…  «·ﬂ»—Ì  " & UnitName
 Else
 
     Reportname = "Inventory\sales65.rpt"
        StrTitle = "Sales Retuen Totals By Items big Unit   " & UnitName
 End If
 
 
   ElseIf newReport = 965 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales965.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„»Ì⁄«    ··⁄„·«¡  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales965.rpt"
        StrTitle = "Purcahse Retuen Totals By Items For Unit   " & UnitName
 End If
 
    ElseIf newReport = 38 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales38.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›    „Õ„·… »«·„’—Ê›«     " & UnitName
Else
      Reportname = "Inventory\sales38E.rpt"
        StrTitle = "Purchase With Its Expenses  " & UnitName
End If

    ElseIf newReport = 932 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales932.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì       "
Else
      Reportname = "Inventory\sales932E.rpt"
        StrTitle = "Total Purcahse By Items"
End If


    ElseIf newReport = 963 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales963.rpt"
        StrTitle = " ﬁ—Ì—  „—œÊœ«  «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì       "
Else
      Reportname = "Inventory\sales963E.rpt"
        StrTitle = "Total Purcahse By Items"
End If




    ElseIf newReport = 11 Then
        If SystemOptions.UserInterface = ArabicInterface Then
   isGrouping = MsgBox("Â·  —Ìœ  Ã„Ì⁄ ⁄·Ì „” ÊÌ «·⁄„Ì·", vbInformation + vbYesNo)
   Else
   isGrouping = MsgBox("Do you want group by cutomer y/n", vbInformation + vbYesNo)
   End If
   
   
      If SystemOptions.UserInterface = ArabicInterface Then
       If isGrouping = vbYes Then
        Reportname = "Inventory\sales6.rpt"
       Else
       Reportname = "Inventory\sales6WithoutGrouping.rpt"
       End If
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  Õ·Ì·Ì " & UnitName
      Else
      If isGrouping = vbYes Then
      Reportname = "Inventory\sales6E.rpt"
      Else
      Reportname = "Inventory\sales6EWithoutGrouping.rpt"
      End If
        StrTitle = "Sales Report Details By Customers for Unit " & UnitName
      End If
      
        
    ElseIf newReport = 911 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
   isGrouping = MsgBox("Â·  —Ìœ  Ã„Ì⁄ ⁄·Ì „” ÊÌ «·⁄„Ì·", vbInformation + vbYesNo)
   Else
   isGrouping = MsgBox("Do you want group by cutomer y/n", vbInformation + vbYesNo)
   End If
   
    
          If SystemOptions.UserInterface = ArabicInterface Then
                                If isGrouping = vbYes Then
                                        Reportname = "Inventory\sales911.rpt"
                                Else
                                Reportname = "Inventory\sales911WithoutGrouping.rpt"
                                End If
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  Õ·Ì·Ì "
Else
  
  If isGrouping = vbYes Then
  Reportname = "Inventory\sales911E.rpt"
  Else
  Reportname = "Inventory\sales911EWithoutGrouping.rpt"
  End If
  
StrTitle = "Sales Report Details By Customers " & UnitName
End If

    ElseIf newReport = 30 Then
                  If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales30.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  Õ·Ì·Ì " & UnitName
Else

        Reportname = "Inventory\sales30E.rpt"
        StrTitle = "Detailed Purchases By Supplier For Unit: " & UnitName
End If


    ElseIf newReport = 930 Then
              If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales930.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰   Õ·Ì·Ì  "
Else
   Reportname = "Inventory\sales930E.rpt"
        StrTitle = "Purchase  Report Detailed By Supplier "

End If
    ElseIf newReport = 15 Then
        Reportname = "Inventory\sales7.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«   ÿ»ﬁ« ·ÊÕœ«  «·»Ì⁄ "
    ElseIf newReport = 16 Then
        Reportname = "Inventory\sales8.rpt"
        StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 160 Then
  Dim X As Integer
  
  StrTitle = reportString
'      If salesflag = "SalesPrice1" Then
'    Reportname = "Inventory\sales160-1sales.rpt"
'     ElseIf salesflag = "wholePrice'" Then
'            Reportname = "Inventory\sales160-1wholePrice.rpt"
'
            
            
            If salesflag = "SalesPrice1" Or salesflag = "wholePrice'" Then
            
                   StrTitle = reportString
       If SystemOptions.UserInterface = ArabicInterface Then

    
                X = MsgBox("›—“ »«·„Ã„Ê⁄Â", vbInformation + vbYesNoCancel)

                        If X = vbYes Then
                                   If salesflag = "SalesPrice1" Then
                                        Reportname = "Inventory\sales160-1sales.rpt"
                                    ElseIf salesflag = "wholePrice'" Then
                                        Reportname = "Inventory\sales160-1wholePrice.rpt"
                                    Else
                                        Reportname = "Inventory\sales160-1.rpt"
                                    End If
                                    
                        ElseIf X = vbNo Then
                                If salesflag = "SalesPrice1" Then
                                    Reportname = "Inventory\sales160-1salesTotal.rpt"
                                ElseIf salesflag = "wholePrice'" Then
                                    Reportname = "Inventory\sales160-1wholePriceTotal.rpt"
                                Else
                                
                                End If
                        Else
                                   Exit Function
                        End If
                        
        
    Else
                    X = MsgBox("Grouping by item Group  ", vbInformation + vbYesNoCancel)
                            If X = vbYes Then
                                       Reportname = "Inventory\sales160E-1.rpt"
                            ElseIf X = vbNo Then
                                      Reportname = "Inventory\sales160E.rpt"
                            Else
                                       Exit Function
                            End If
                    
    
   '  ReportName = "Inventory\sales160E.rpt"
    End If

            
            
            
Else
    
    
   If SystemOptions.HideCost = False Then
    
    If SystemOptions.UserInterface = ArabicInterface Then

    
                X = MsgBox("›—“ »«·„Ã„Ê⁄Â", vbInformation + vbYesNoCancel)

                        If X = vbYes Then
                                   Reportname = "Inventory\sales160-1.rpt"
                        ElseIf X = vbNo Then
                                  Reportname = "Inventory\sales160.rpt"
                        Else
                                   Exit Function
                        End If
                        
        
    Else
                    X = MsgBox("Grouping by item Group  ", vbInformation + vbYesNoCancel)
                            If X = vbYes Then
                                       Reportname = "Inventory\sales160E-1.rpt"
                            ElseIf X = vbNo Then
                                      Reportname = "Inventory\sales160E.rpt"
                            Else
                                       Exit Function
                            End If
                    
    
   '  ReportName = "Inventory\sales160E.rpt"
    End If
    
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales160hide.rpt"
    Else
     Reportname = "Inventory\sales160Ehide.rpt"
    End If
    End If
    
    
    
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString
    End If
    ElseIf newReport = 1600 Then

                 If SystemOptions.UserInterface = ArabicInterface Then
        
                Reportname = "Inventory\sales1600.rpt"
                Else
                Reportname = "Inventory\sales1600e.rpt"
                End If
 
        
        'StrTitle = " «·”«·» ›ﬁÿ « ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 16000 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales16000.rpt"
Else
        Reportname = "Inventory\sales16000e.rpt"
End If


    ElseIf newReport = 2092014 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "REPORTS NEW\IssueItemsBystore.rpt"
Else
        Reportname = "Inventory\IssueItemsBystore.rpt"
End If

        'StrTitle = "    œÊ—«‰ «·«’‰«› „⁄œ· "
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString


   ElseIf newReport = 28116 Then
   
     If SystemOptions.UserInterface = ArabicInterface Then
     StrTitle = "„Êﬁ› «·„Œ“Ê‰ ÿ»ﬁ« ·—ﬁ„ «··Êÿ "
                 Reportname = "Inventory\IventoryStockWithLOT.rpt"

     Else
                 Reportname = "Inventory\IventoryStockWithLOTe.rpt"

     End If
     
      ElseIf newReport = 281161 Then
   
     If SystemOptions.UserInterface = ArabicInterface Then
     StrTitle = "„Êﬁ› «·„Œ“Ê‰ ÿ»ﬁ«  · «—Œ «·«‰ Â«¡ "
                 Reportname = "Inventory\IventoryStockWithExpDate.rpt"

     Else
                 Reportname = "Inventory\IventoryStockWithExpDateE.rpt"

     End If
    ElseIf newReport = 17 Then
If SystemOptions.UserInterface = ArabicInterface Then
        If showvalue = False Then
            Reportname = "Inventory\sales9.rpt"
        Else
            Reportname = "Inventory\sales9value.rpt"
        End If
Else

       If showvalue = False Then
            Reportname = "Inventory\sales9e.rpt"
        Else
            Reportname = "Inventory\sales9valuee.rpt"
        End If
End If
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 181 Then

    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits.rpt"
    Else
    Reportname = "New Reports\GardByUnitsE.rpt"
    End If


      ElseIf newReport = 18 Then
If ViewCost = 1 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits2.rpt"
    Else
    Reportname = "New Reports\GardByUnits2E.rpt"
    End If
 Else
 
     If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits1.rpt"
    Else
    Reportname = "New Reports\GardByUnits1E.rpt"
    End If
 End If
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 19 Then
        Reportname = "New Reports\PerimimPayed.rpt"
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 20 Then
        Reportname = "New Reports\PerimimPayedFromCustomer.rpt"
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 24 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales24.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
       Else
        Reportname = "Inventory\sales24eE.rpt"
        StrTitle = " Totals Sales By Customer Monthly by Unit  " & UnitName
 
        End If
        ElseIf newReport = 57 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales57.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
       Else
        Reportname = "Inventory\sales57EE.rpt"
        StrTitle = " Totals Sales By Customer Monthly by Unit  " & UnitName
 
        End If
 
    ElseIf newReport = 924 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales924.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   "
   Else
      Reportname = "Inventory\sales924e.rpt"
        StrTitle = " Totals Sales By Customer Monthly " & UnitName
   End If
   ElseIf newReport = 957 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales957.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   "
   Else
      Reportname = "Inventory\sales957e.rpt"
        StrTitle = " Totals Sales By Customer Monthly " & UnitName
   End If
   
    ElseIf newReport = 35 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales35.rpt"
        StrTitle = " ﬁ—Ì—   «Ã„«·Ì «·„»Ì⁄«   ··⁄„·«¡   ⁄‰ › —…   " & UnitName
 Else
 Reportname = "Inventory\sales35E.rpt"
        StrTitle = "Total Customer Sales   " & UnitName
 End If
 
    ElseIf newReport = 37 Then
            If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales37.rpt"
        StrTitle = " ﬁ—Ì—   «Ã„«·Ì «·„‘ —Ì«  „‰ «·„Ê—œÌ‰    ⁄‰ › —…   "
Else
 Reportname = "Inventory\sales37E.rpt"
        StrTitle = "Totals Purchase By Supplier   "

End If
    ElseIf newReport = 31 Then
                If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales31.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«   ‘Â—Ì« »«·ÊÕœ… " & UnitName
 Else
 
        Reportname = "Inventory\sales31E.rpt"
        StrTitle = "Purchases Totals By Employee For Unit " & UnitName
 End If
    ElseIf newReport = 931 Then
            If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales931.rpt"
        StrTitle = "  ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«   ‘Â—Ì«    "
 Else
    Reportname = "Inventory\sales931E.rpt"
        StrTitle = "  Totals Purchase By Supplier   Monthly    "
 End If
    
    ElseIf newReport = 25 Then
       
       
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales25.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales25E.rpt"
        StrTitle = " MonthlyTotal Sales  by Units: " & UnitName
     End If
     
    ElseIf newReport = 925 Then
           If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales925.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
 Else
  Reportname = "Inventory\sales925E.rpt"
        StrTitle = "Monthly Total Sales" & UnitName
 
 End If
'************************************************************************
    ElseIf newReport = 39 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales39.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales39E.rpt"
        StrTitle = " MonthlyTotal Sales  by Units: " & UnitName
     End If
     
    ElseIf newReport = 939 Then
                   If SystemOptions.UserInterface = ArabicInterface Then
        
                Reportname = "Inventory\sales939.rpt"
                StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
                Else
                 Reportname = "Inventory\sales939E.rpt"
                       StrTitle = "Monthly Total Sales" & UnitName
                
                End If

'***********************************************************
    ElseIf newReport = 45 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales45.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«› ÿ»ﬁ« ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales45E.rpt"
        StrTitle = " MonthlyTotal Sales  by Sales Person Units: " & UnitName
     End If
     
     
         ElseIf newReport = 945 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales945.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  ÿ»ﬁ« ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«      " & UnitName
     Else
        Reportname = "Inventory\sales945E.rpt"
        StrTitle = " MonthlyTotal Sales  by Sales Person  : " & UnitName
     End If
     
     
    ElseIf newReport = 36 Then
    Dim checkgroup As Integer
    If SystemOptions.UserInterface = ArabicInterface Then
    checkgroup = MsgBox(" Ã„Ì⁄ ⁄·Ì „” ÊÌ «·„‰œÊ» ", vbInformation + vbYesNo)
    Else
    checkgroup = MsgBox("Group by employee ", vbInformation + vbYesNo)
    End If
    
    If checkgroup = vbYes Then
    
                      If SystemOptions.UserInterface = ArabicInterface Then
                      Reportname = "Inventory\sales36Grouped.rpt"
                      StrTitle = " ﬁ—Ì— «Ã„«·Ì «·„»Ì⁄«    ··«’‰«› »«·ÊÕœ…" & UnitName
                Else
                  Reportname = "Inventory\sales36GroupeE.rpt"
                      StrTitle = "Sales Total by Unit :" & UnitName
                End If
                
    Else
                  If SystemOptions.UserInterface = ArabicInterface Then
                      Reportname = "Inventory\sales36.rpt"
                      StrTitle = " ﬁ—Ì— «Ã„«·Ì «·„»Ì⁄«    ··«’‰«› »«·ÊÕœ…" & UnitName
                Else
                  Reportname = "Inventory\sales36E.rpt"
                      StrTitle = "Sales Total by Unit :" & UnitName
                End If
  End If
  
  ElseIf newReport = 600 Then
  
       
    X = MsgBox("⁄—÷ »‘ﬂ· «›ﬁÌ", vbInformation + vbYesNo)
    
    If X = vbNo Then
    
                If SystemOptions.UserInterface = ArabicInterface Then
                          If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600.rpt"
                           Else
                           Reportname = "Inventory\sales600NG.rpt"
                           End If
                 Else
                            If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600E.rpt"
                           Else
                           Reportname = "Inventory\sales600NGE.rpt"
                           End If
                 
                 
                 End If
   Else
                   If SystemOptions.UserInterface = ArabicInterface Then
                          If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600land.rpt"
                           Else
                           Reportname = "Inventory\sales600NGland.rpt"
                           End If
                 Else
                            If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600ElandE.rpt"
                           Else
                           Reportname = "Inventory\sales600NGElandE.rpt"
                           End If
                 
                 
                 End If

   
   End If
   
 If SystemOptions.UserInterface = ArabicInterface Then
                      StrTitle = " ﬁ—Ì—«—’œ… «·⁄„·«¡ " & UnitName
  Else
               StrTitle = "Customer Balance" & UnitName
  End If
  
                If reportString <> "" Then
                StrTitle = StrTitle & "    " & reportString
                End If
                
  
    ElseIf newReport = 936 Then
    
    
        If SystemOptions.UserInterface = ArabicInterface Then
    checkgroup = MsgBox(" Ã„Ì⁄ ⁄·Ì „” ÊÌ «·„‰œÊ» ", vbInformation + vbYesNo)
    Else
    checkgroup = MsgBox("Group by employee ", vbInformation + vbYesNo)
    End If
    
    If checkgroup = vbYes Then
                If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales936grouped.rpt"
        StrTitle = " ﬁ—Ì—  «Ã„«·Ì «·„»Ì⁄«    ··«’‰«›   "
 Else
     Reportname = "Inventory\sales936groupedE.rpt"
        StrTitle = "Sales Total  "
 End If
 
 
    
   Else
                      If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales936.rpt"
        StrTitle = " ﬁ—Ì—  «Ã„«·Ì «·„»Ì⁄«    ··«’‰«›   "
 Else
     Reportname = "Inventory\sales936E.rpt"
        StrTitle = "Sales Total  "
 End If
 
 
 End If
    ElseIf newReport = 33 Then
  If SystemOptions.UserInterface = ArabicInterface Then
   Reportname = "Inventory\sales33.rpt"
                  StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«  ‘Â—Ì«   "
             
 Else
     Reportname = "Inventory\sales33E.rpt"
 StrTitle = "Monthly Purchase Report  "
 End If
 
    ElseIf newReport = 933 Then
        Reportname = "Inventory\sales933.rpt"
                      If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«  ‘Â—Ì«   " & UnitName
                    Else
                    StrTitle = "Monthly Purchase Report" & UnitName
                    End If
        
        'StrTitle = " ﬁ—Ì—      «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì   ‘Â—Ì«   " & Unitname
 
    ElseIf newReport = 26 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                       Reportname = "Inventory\sales26.rpt"
                       StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  Õ·Ì·Ì    »«·ÊÕœ…" & UnitName
                Else
                  Reportname = "Inventory\sales26e.rpt"
                       StrTitle = "Sales For Sales Person Detailed By Unit" & UnitName
                End If
 
    ElseIf newReport = 926 Then
    
                        If SystemOptions.UserInterface = ArabicInterface Then
    Reportname = "Inventory\sales926.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  Õ·Ì·Ì     "
  Else
     Reportname = "Inventory\sales926E.rpt"
        StrTitle = "Sales For Sales Person    "
  End If
  
  
X = MsgBox("›—“ »«·„Ã„Ê⁄Â", vbInformation + vbYesNo)
If X = vbYes Then

   If SystemOptions.UserInterface = ArabicInterface Then
    Reportname = "Inventory\sales926Grouos.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  Õ·Ì·Ì     "
  Else
     Reportname = "Inventory\sales926GrouosE.rpt"
        StrTitle = "Sales For Sales Person    "
  End If
 

End If


    ElseIf newReport = 27 Then
      If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales27.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
    Else
    
      Reportname = "Inventory\sales27e.rpt"
        StrTitle = "Sales For Sales Person  By Unit" & UnitName
    End If
    
    ElseIf newReport = 927 Then
                  If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales27.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
  Else
        Reportname = "Inventory\sales27E.rpt"
        StrTitle = "Sales For Sales Person    " & UnitName
  End If
    ElseIf newReport = 28 Then
              If SystemOptions.UserInterface = ArabicInterface Then
                  Reportname = "Inventory\sales28.rpt"
                  StrTitle = " ﬁ—»—  Õ’Ì·«  «·„‰«œÌ»  Õ·Ì·Ì" & UnitName
            Else
              Reportname = "Inventory\sales28E.rpt"
                  StrTitle = "Collection Reports For Sales Person" & UnitName
            End If
            
    ElseIf newReport = 29 Then
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
  
                    Reportname = "Inventory\sales29.rpt"
                    StrTitle = "  ﬁ—Ì—«Ã„«·Ì   Õ’Ì·«  «·„‰«œÌ» ‘Â—Ì«    "
            Else
                                Reportname = "Inventory\sales29E.rpt"
                    StrTitle = " Sales Person Collection Monthly  "

            End If
            
            
                ElseIf newReport = 78 Then
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
  
                    Reportname = "Inventory\sales78.rpt"
                    StrTitle = "  ﬁ—Ì—«Ã„«·Ì   Õ’Ì·«  «·„‰«œÌ»      "
            Else
                                Reportname = "Inventory\sales29E.rpt"
                    StrTitle = " Sales Person Collection Monthly  "

            End If
            
    ElseIf newReport = 44 Then
                    
            If SystemOptions.UserInterface = ArabicInterface Then
  
                    Reportname = "Inventory\serialProfit.rpt"
              '      StrTitle = "  ﬁ—Ì—«Ã„«·Ì   Õ’Ì·«  «·„‰«œÌ» ‘Â—Ì«    "
            Else
                                Reportname = "Inventory\serialProfitE.rpt"
              '      StrTitle = " Sales Person Collection Monthly  "

            End If
            
            'serialProfit

    End If

If branchname <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··›—⁄ :" & branchname
                Else
                          StrTitle = StrTitle & " Branch :" & branchname
                End If

End If


If GroupName <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··„Ã„Ê⁄Â :" & GroupName
                Else
                          StrTitle = StrTitle & " Group :" & GroupName
                End If

End If


If CustomerLocation <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··„Êﬁ⁄ :" & CustomerLocation
                Else
                          StrTitle = StrTitle & " Location :" & CustomerLocation
                End If

End If

    If mReportName <> "" Then Reportname = mReportName

    StrFileName = App.path & "\Reports\" & Reportname

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)

        If BolShowNotes = False Then
            xReport.Database.SetDataSource rs
        Else
        
            xReport.Database.Tables(1).SetDataSource rs
            xReport.Database.Tables(2).SetDataSource RsNotes
        End If

        '---------------------------------------------------
        If Reportname = "ReportSalingTimeNotTax.rpt" Then
            xReport.OpenSubreport("SubReport1").Database.SetDataSource rs
            xReport.OpenSubreport("SubRptCustomerTrans.rpt").Database.SetDataSource rs
        End If
    
        Set cOptions = New ClsCompanyInfo
  If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
   Else
   xReport.ParameterFields(1).AddCurrentValue cOptions.EngCompanyName
   End If
  If newReport = 8 Then
    xReport.ParameterFields(6).AddCurrentValue CStr(SystemOptions.Count_ACCOUNT_digit)
  End If
  
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
    
        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If
    
        If Reportname = "ReportSalingTimeNotTax.rpt" Then

            For i = 1 To xReport.ParameterFields.count
                xReport.OpenSubreport("SubReport1").ParameterFields(i).AddCurrentValue xReport.ParameterFields(i).GetNthCurrentValue(1)
            Next i

            For i = 1 To xReport.ParameterFields.count
                xReport.OpenSubreport("SubRptCustomerTrans.rpt").ParameterFields(i).AddCurrentValue xReport.ParameterFields(i).GetNthCurrentValue(1)
            Next i

        End If

        Screen.MousePointer = vbDefault
    End If

    xReport.reporttitle = StrTitle

   

    If newReport = 18 Or newReport = 181 Then
        'x = MsgBox("Â·  —Ìœ ⁄—÷ „ Ê”ÿ «· ﬂ·›…  ‰⁄„ «„ ·« ", vbCritical + vbYesNo)
        'If x = vbYes Then
        xReport.ParameterFields(6).AddCurrentValue ViewCost
        'Else
        'xReport.ParameterFields(6).AddCurrentValue "0"
        'End If
    End If



    If newReport = 600 Then
      '    xReport.ParameterFields(6).AddCurrentValue ViewCost
          xReport.ParameterFields(7).AddCurrentValue ShowZero
    End If
    
    
    If newReport = 1001 Or newReport = 1002 Then
        If ReturnPurcahse <> 0 Then
            xReport.ParameterFields(7).AddCurrentValue "«· ﬁ—Ì— ·« ÌÕ ÊÌ ”‰œ«  ’—› „—œÊœ«  «·„‘ —Ì« "
        End If

        xReport.ParameterFields(7).AddCurrentValue ""

    End If
    
    If mTypeDebitCredit = 1 Then
        xReport.RecordSelectionFormula = "{@netBalanceCredit} = 0"
    ElseIf mTypeDebitCredit = 2 Then
        xReport.RecordSelectionFormula = "{@NetBalanceDebit} = 0"
    End If
    

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , StrFileName, , StrSQL
    Set CViewer = Nothing
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Msg = Msg + CHR(13) & "Err.Description:" & Err.description
    Msg = Msg + CHR(13) & "Err.Source:" & Err.Source
    Msg = Msg + CHR(13) & "Err.Number:" & Err.Number
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function





 
   Public Function ShowSallingTime05042019(StrSQL As String, _
                                Optional DateFrom As Variant = Null, _
                                Optional DateTo As Variant = Null, _
                                Optional BolShowNotes As Boolean = False, _
                                Optional newReport As Variant, _
                                Optional PaymentType As Integer = 2, _
                                Optional reportString As String, _
                                Optional UnitName As String, _
                                Optional ViewCost As String, _
                                Optional reporttitle As String = "", _
                                Optional ReturnPurcahse As Integer = 0, _
                                Optional showvalue As Boolean = False, Optional branchname As String = "", Optional ShowZero As String, Optional Departement As String, Optional GroupName As String, Optional ShowGrouping As Integer = 1, Optional CustomerLocation As String, Optional ByVal mTypeDebitCredit As Integer = 0)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim Reportname As String
    Dim ShowTax As Boolean
    Dim RsNotes As ADODB.Recordset
    Dim MySQL As String
    Dim StrTitle As String
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim i As Integer
    Dim StrFileName As String

    On Error GoTo ErrTrap

     If StrSQL = "" Then
        Exit Function
       End If

    'Œ«’ »÷—Ì»…«·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    ShowTax = False

    If ShowTax = True Then
        Reportname = "ReportSalingTime.rpt"
    ElseIf BolShowNotes = False Then
        Reportname = "ReportSalingTimeNotTax.rpt"
    ElseIf BolShowNotes = True Then
        Reportname = "ReportSalingTimeNotTaxDebit.rpt"
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«  ÊÃœ »Ì«‰«  ›Ì Â–Â «·› —…"
 Else
 Msg = "No Data in this Period"
 End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If BolShowNotes = True Then
        StrTitle = " ﬁ—Ì— »«·„»Ì⁄«  «·√Ã·… Ê«· Õ’Ì·«  «·ﬁ«∆„… ⁄·ÌÂ«"
        Set RsNotes = New ADODB.Recordset
        MySQL = "Select * From QryPostedTransNotes Where NoteType=4 OR NoteType=9"
        RsNotes.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrTitle = " ﬁ—Ì— »«·„»Ì⁄« "
    Else
    StrTitle = "Sales report "
    End If
        '    If Not (IsNull(DateFrom)) Then
        '        StrTitle = StrTitle & " ›Ï «·› —… „‰ " & DisplayDate(CDate(DateFrom)) & "¯"
        '    End If
        '    If Not (IsNull(DateTo)) Then
        '        StrTitle = StrTitle & " ≈·Ï " & DisplayDate(CDate(DateTo))
        '    End If
        '    StrTitle = StrTitle & "¯"
    End If

    If newReport = 1 Then
        Reportname = "Inventory\sales2.rpt"
    ElseIf newReport = 2 Then
        Reportname = "Inventory\salesByBranch.rpt"
    ElseIf newReport = 3 Then
        Reportname = "Inventory\salesByBranchYearly.rpt"
    ElseIf newReport = 4 Then
        Reportname = "Inventory\sales3.rpt"
        StrTitle = " ﬁ—Ì— „»Ì⁄«  ﬂ· «·„Œ«“‰ √›ﬁÌ"
    ElseIf newReport = 5 Then
        Reportname = "Inventory\ItemReport.rpt"
        StrTitle = " ﬁ—Ì— «·«’‰«›"
    ElseIf newReport = 6 Then
        Reportname = "Inventory\Gard4.rpt"
        StrTitle = " ﬁ—Ì—„ «»⁄Â  €Ì— «”⁄«— «·«’‰«›"
    ElseIf newReport = 7 Then
        Reportname = "Inventory\sales4.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    «·«’‰«›"

    ElseIf newReport = 8 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "CommercialTransactions\CommercialTransactions.rpt"
Else
 Reportname = "CommercialTransactions\CommercialTransactionse.rpt"
End If



If Departement <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
StrTitle = "„—«Ã⁄Â «·„⁄«„·«  «· Ã«—Ì… "
Else
StrTitle = "Commercial Transaction Auditing"
End If
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··«œ«—… :" & Departement
                Else
                          StrTitle = StrTitle & " Departement :" & Departement
                End If

End If



    ElseIf newReport = 845 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "CommercialTransactions\CommercialTransactions1.rpt"
     Else
     Reportname = "CommercialTransactions\CommercialTransactions1e.rpt"
     End If
        
        If Departement <> "" Then
StrTitle = "„—«Ã⁄Â «·„⁄«„·«  «· Ã«—Ì… "
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··«œ«—… :" & Departement
                Else
                          StrTitle = StrTitle & " Departement :" & Departement
                End If

End If


    ElseIf newReport = 9 Then
    If UnitName <> "" Then
     StrTitle = " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì… ··„‰œÊ» " & CHR(13) & UnitName
    Else
     StrTitle = " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì…   "
     
    End If
    
        Reportname = "AccountsTransactions\AccountsTransactions.rpt"
 
    ElseIf newReport = 10 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales5.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
    Else
    Reportname = "Inventory\sales5E.rpt"
        StrTitle = "Sales Report Detailed By Unit :" & UnitName
    
    End If
    
    
    ElseIf newReport = 910 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales910.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  Õ·Ì·Ì       " & UnitName
Else

        Reportname = "Inventory\sales910E.rpt"
        StrTitle = "Sales Report Detailed by Items  "

End If

    ElseIf newReport = 1001 Then
        Reportname = "Inventory\sales1001.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì »«·ÊÕœ… " & UnitName
    ElseIf newReport = 1002 Then
        Reportname = "Inventory\sales1002.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì   "
    ElseIf newReport = 1003 Then
        Reportname = "Inventory\sales1003.rpt"
        StrTitle = " ﬁ—Ì— ”‰œ«    " & reporttitle & "  ··«’‰«›  Õ·Ì·Ì   "


    ElseIf newReport = 32 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales32.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales32E.rpt"
        StrTitle = "Purcahse Totals By Items For Unit   " & UnitName
 End If
    
    
    ElseIf newReport = 63 Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales63.rpt"
        StrTitle = " ﬁ—Ì— „—œÊœ«   «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì  »«·ÊÕœ…    " & UnitName
 Else
 
     Reportname = "Inventory\sales63E.rpt"
        StrTitle = "Purcahse Retuen Totals By Items For Unit   " & UnitName
 End If
 
 
    ElseIf newReport = 38 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales38.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›    „Õ„·… »«·„’—Ê›«     " & UnitName
Else
      Reportname = "Inventory\sales38E.rpt"
        StrTitle = "Purchase With Its Expenses  " & UnitName
End If

    ElseIf newReport = 932 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales932.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì       "
Else
      Reportname = "Inventory\sales932E.rpt"
        StrTitle = "Total Purcahse By Items"
End If


    ElseIf newReport = 963 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales963.rpt"
        StrTitle = " ﬁ—Ì—  „—œÊœ«  «·„‘ —Ì«    ··«’‰«›  Õ·Ì·Ì       "
Else
      Reportname = "Inventory\sales963E.rpt"
        StrTitle = "Total Purcahse By Items"
End If




    ElseIf newReport = 11 Then
      If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales6.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  Õ·Ì·Ì " & UnitName
      Else
      Reportname = "Inventory\sales6E.rpt"
        StrTitle = "Sales Report Details By Customers for Unit " & UnitName
      End If
      
        
    ElseIf newReport = 911 Then
          If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales911.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  Õ·Ì·Ì "
Else
  Reportname = "Inventory\sales911E.rpt"
StrTitle = "Sales Report Details By Customers " & UnitName
End If

    ElseIf newReport = 30 Then
                  If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales30.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  Õ·Ì·Ì " & UnitName
Else

        Reportname = "Inventory\sales30E.rpt"
        StrTitle = "Detailed Purchases By Supplier For Unit: " & UnitName
End If


    ElseIf newReport = 930 Then
              If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales930.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰   Õ·Ì·Ì  "
Else
   Reportname = "Inventory\sales930E.rpt"
        StrTitle = "Purchase  Report Detailed By Supplier "

End If
    ElseIf newReport = 15 Then
        Reportname = "Inventory\sales7.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«   ÿ»ﬁ« ·ÊÕœ«  «·»Ì⁄ "
    ElseIf newReport = 16 Then
        Reportname = "Inventory\sales8.rpt"
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 160 Then
    
   If SystemOptions.HideCost = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales160.rpt"
    Else
     Reportname = "Inventory\sales160E.rpt"
    End If
    
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales160hide.rpt"
    Else
     Reportname = "Inventory\sales160Ehide.rpt"
    End If
    End If
    
    
    
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 1600 Then

                 If SystemOptions.UserInterface = ArabicInterface Then
        
                Reportname = "Inventory\sales1600.rpt"
                Else
                Reportname = "Inventory\sales1600e.rpt"
                End If
 
        
        'StrTitle = " «·”«·» ›ﬁÿ « ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 16000 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales16000.rpt"
Else
        Reportname = "Inventory\sales16000e.rpt"
End If


    ElseIf newReport = 2092014 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "REPORTS NEW\IssueItemsBystore.rpt"
Else
        Reportname = "Inventory\IssueItemsBystore.rpt"
End If

        'StrTitle = "    œÊ—«‰ «·«’‰«› „⁄œ· "
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString


   ElseIf newReport = 28116 Then
   
     If SystemOptions.UserInterface = ArabicInterface Then
     StrTitle = "„Êﬁ› «·„Œ“Ê‰ ÿ»ﬁ« ·—ﬁ„ «··Êÿ "
                 Reportname = "Inventory\IventoryStockWithLOT.rpt"

     Else
                 Reportname = "Inventory\IventoryStockWithLOTe.rpt"

     End If
     
      ElseIf newReport = 281161 Then
   
     If SystemOptions.UserInterface = ArabicInterface Then
     StrTitle = "„Êﬁ› «·„Œ“Ê‰ ÿ»ﬁ«  · «—Œ «·«‰ Â«¡ "
                 Reportname = "Inventory\IventoryStockWithExpDate.rpt"

     Else
                 Reportname = "Inventory\IventoryStockWithExpDateE.rpt"

     End If
    ElseIf newReport = 17 Then
If SystemOptions.UserInterface = ArabicInterface Then
        If showvalue = False Then
            Reportname = "Inventory\sales9.rpt"
        Else
            Reportname = "Inventory\sales9value.rpt"
        End If
Else

       If showvalue = False Then
            Reportname = "Inventory\sales9e.rpt"
        Else
            Reportname = "Inventory\sales9valuee.rpt"
        End If
End If
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 181 Then

    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits.rpt"
    Else
    Reportname = "New Reports\GardByUnitsE.rpt"
    End If


      ElseIf newReport = 18 Then
If ViewCost = 1 Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits2.rpt"
    Else
    Reportname = "New Reports\GardByUnits2E.rpt"
    End If
 Else
 
     If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "New Reports\GardByUnits1.rpt"
    Else
    Reportname = "New Reports\GardByUnits1E.rpt"
    End If
 End If
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 19 Then
        Reportname = "New Reports\PerimimPayed.rpt"
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 20 Then
        Reportname = "New Reports\PerimimPayedFromCustomer.rpt"
        'StrTitle = " ﬁ—Ì—     „Êﬁ› «·„Œ“Ê‰ «·Õ«·Ì"
        StrTitle = ""
        StrTitle = StrTitle & CHR(13) & reportString

    ElseIf newReport = 24 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales24.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
       Else
        Reportname = "Inventory\sales24eE.rpt"
        StrTitle = " Totals Sales By Customer Monthly by Unit  " & UnitName
 
        End If
        ElseIf newReport = 57 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales57.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
       Else
        Reportname = "Inventory\sales57EE.rpt"
        StrTitle = " Totals Sales By Customer Monthly by Unit  " & UnitName
 
        End If
 
    ElseIf newReport = 924 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales924.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   "
   Else
      Reportname = "Inventory\sales924e.rpt"
        StrTitle = " Totals Sales By Customer Monthly " & UnitName
   End If
   ElseIf newReport = 957 Then
        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales957.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··⁄„·«¡  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   "
   Else
      Reportname = "Inventory\sales957e.rpt"
        StrTitle = " Totals Sales By Customer Monthly " & UnitName
   End If
   
    ElseIf newReport = 35 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales35.rpt"
        StrTitle = " ﬁ—Ì—   «Ã„«·Ì «·„»Ì⁄«   ··⁄„·«¡   ⁄‰ › —…   " & UnitName
 Else
 Reportname = "Inventory\sales35E.rpt"
        StrTitle = "Total Customer Sales   " & UnitName
 End If
 
    ElseIf newReport = 37 Then
            If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales37.rpt"
        StrTitle = " ﬁ—Ì—   «Ã„«·Ì «·„‘ —Ì«  „‰ «·„Ê—œÌ‰    ⁄‰ › —…   "
Else
 Reportname = "Inventory\sales37E.rpt"
        StrTitle = "Totals Purchase By Supplier   "

End If
    ElseIf newReport = 31 Then
                If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales31.rpt"
        StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«   ‘Â—Ì« »«·ÊÕœ… " & UnitName
 Else
 
        Reportname = "Inventory\sales31E.rpt"
        StrTitle = "Purchases Totals By Employee For Unit " & UnitName
 End If
    ElseIf newReport = 931 Then
            If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales931.rpt"
        StrTitle = "  ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«   ‘Â—Ì«    "
 Else
    Reportname = "Inventory\sales931E.rpt"
        StrTitle = "  Totals Purchase By Supplier   Monthly    "
 End If
    
    ElseIf newReport = 25 Then
       
       
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales25.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales25E.rpt"
        StrTitle = " MonthlyTotal Sales  by Units: " & UnitName
     End If
     
    ElseIf newReport = 925 Then
           If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales925.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
 Else
  Reportname = "Inventory\sales925E.rpt"
        StrTitle = "Monthly Total Sales" & UnitName
 
 End If
'************************************************************************
    ElseIf newReport = 39 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales39.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales39E.rpt"
        StrTitle = " MonthlyTotal Sales  by Units: " & UnitName
     End If
     
    ElseIf newReport = 939 Then
                   If SystemOptions.UserInterface = ArabicInterface Then
        
                Reportname = "Inventory\sales939.rpt"
                StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
                Else
                 Reportname = "Inventory\sales939E.rpt"
                       StrTitle = "Monthly Total Sales" & UnitName
                
                End If

'***********************************************************
    ElseIf newReport = 45 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales45.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«› ÿ»ﬁ« ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
     Else
        Reportname = "Inventory\sales45E.rpt"
        StrTitle = " MonthlyTotal Sales  by Sales Person Units: " & UnitName
     End If
     
     
         ElseIf newReport = 945 Then
          
       If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales945.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··«’‰«›  ÿ»ﬁ« ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«      " & UnitName
     Else
        Reportname = "Inventory\sales945E.rpt"
        StrTitle = " MonthlyTotal Sales  by Sales Person  : " & UnitName
     End If
     
     
    ElseIf newReport = 36 Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                      Reportname = "Inventory\sales36.rpt"
                      StrTitle = " ﬁ—Ì— «Ã„«·Ì «·„»Ì⁄«    ··«’‰«› »«·ÊÕœ…" & UnitName
                Else
                  Reportname = "Inventory\sales36E.rpt"
                      StrTitle = "Sales Total by Unit :" & UnitName
                End If
  
  ElseIf newReport = 600 Then
  
      Dim X As Integer
    X = MsgBox("⁄—÷ »‘ﬂ· «›ﬁÌ", vbInformation + vbYesNo)
    
    If X = vbNo Then
    
                If SystemOptions.UserInterface = ArabicInterface Then
                          If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600.rpt"
                           Else
                           Reportname = "Inventory\sales600NG.rpt"
                           End If
                 Else
                            If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600E.rpt"
                           Else
                           Reportname = "Inventory\sales600NGE.rpt"
                           End If
                 
                 
                 End If
   Else
                   If SystemOptions.UserInterface = ArabicInterface Then
                          If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600land.rpt"
                           Else
                           Reportname = "Inventory\sales600NGland.rpt"
                           End If
                 Else
                            If ShowGrouping = 1 Then
                           Reportname = "Inventory\sales600ElandE.rpt"
                           Else
                           Reportname = "Inventory\sales600NGElandE.rpt"
                           End If
                 
                 
                 End If

   
   End If
   
 If SystemOptions.UserInterface = ArabicInterface Then
                      StrTitle = " ﬁ—Ì—«—’œ… «·⁄„·«¡ " & UnitName
  Else
               StrTitle = "Customer Balance" & UnitName
  End If
  
                If reportString <> "" Then
                StrTitle = StrTitle & "    " & reportString
                End If
                
  
    ElseIf newReport = 936 Then
                      If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales936.rpt"
        StrTitle = " ﬁ—Ì—  «Ã„«·Ì «·„»Ì⁄«    ··«’‰«›   "
 Else
     Reportname = "Inventory\sales936E.rpt"
        StrTitle = "Sales Total  "
 End If
 
    ElseIf newReport = 33 Then
  If SystemOptions.UserInterface = ArabicInterface Then
   Reportname = "Inventory\sales33.rpt"
                  StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«  ‘Â—Ì«   "
             
 Else
     Reportname = "Inventory\sales33E.rpt"
 StrTitle = "Monthly Purchase Report  "
 End If
 
    ElseIf newReport = 933 Then
        Reportname = "Inventory\sales933.rpt"
                      If SystemOptions.UserInterface = ArabicInterface Then
                    StrTitle = " ﬁ—Ì— «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì «·„‘ —Ì«  ‘Â—Ì«   " & UnitName
                    Else
                    StrTitle = "Monthly Purchase Report" & UnitName
                    End If
        
        'StrTitle = " ﬁ—Ì—      «·„‘ —Ì«    „‰ «·„Ê—œÌ‰  «Ã„«·Ì   ‘Â—Ì«   " & Unitname
 
    ElseIf newReport = 26 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                       Reportname = "Inventory\sales26.rpt"
                       StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  Õ·Ì·Ì    »«·ÊÕœ…" & UnitName
                Else
                  Reportname = "Inventory\sales26e.rpt"
                       StrTitle = "Sales For Sales Person Detailed By Unit" & UnitName
                End If
 
    ElseIf newReport = 926 Then
                        If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales926.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  Õ·Ì·Ì     "
  Else
     Reportname = "Inventory\sales926E.rpt"
        StrTitle = "Sales For Sales Person    "
  End If
  
    ElseIf newReport = 27 Then
      If SystemOptions.UserInterface = ArabicInterface Then
        Reportname = "Inventory\sales27.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì« »«·ÊÕœ… " & UnitName
    Else
    
      Reportname = "Inventory\sales27e.rpt"
        StrTitle = "Sales For Sales Person  By Unit" & UnitName
    End If
    
    ElseIf newReport = 927 Then
                  If SystemOptions.UserInterface = ArabicInterface Then

        Reportname = "Inventory\sales27.rpt"
        StrTitle = " ﬁ—Ì— «·„»Ì⁄«    ··„‰«œÌ»  «Ã„«·Ì «·„»Ì⁄«  ‘Â—Ì«   " & UnitName
  Else
        Reportname = "Inventory\sales27E.rpt"
        StrTitle = "Sales For Sales Person    " & UnitName
  End If
    ElseIf newReport = 28 Then
              If SystemOptions.UserInterface = ArabicInterface Then
                  Reportname = "Inventory\sales28.rpt"
                  StrTitle = " ﬁ—»—  Õ’Ì·«  «·„‰«œÌ»  Õ·Ì·Ì" & UnitName
            Else
              Reportname = "Inventory\sales28E.rpt"
                  StrTitle = "Collection Reports For Sales Person" & UnitName
            End If
            
    ElseIf newReport = 29 Then
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
  
                    Reportname = "Inventory\sales29.rpt"
                    StrTitle = "  ﬁ—Ì—«Ã„«·Ì   Õ’Ì·«  «·„‰«œÌ» ‘Â—Ì«    "
            Else
                                Reportname = "Inventory\sales29E.rpt"
                    StrTitle = " Sales Person Collection Monthly  "

            End If
    ElseIf newReport = 44 Then
                    
            If SystemOptions.UserInterface = ArabicInterface Then
  
                    Reportname = "Inventory\serialProfit.rpt"
              '      StrTitle = "  ﬁ—Ì—«Ã„«·Ì   Õ’Ì·«  «·„‰«œÌ» ‘Â—Ì«    "
            Else
                                Reportname = "Inventory\serialProfitE.rpt"
              '      StrTitle = " Sales Person Collection Monthly  "

            End If
            
            'serialProfit

    End If

If branchname <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··›—⁄ :" & branchname
                Else
                          StrTitle = StrTitle & " Branch :" & branchname
                End If

End If


If GroupName <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··„Ã„Ê⁄Â :" & GroupName
                Else
                          StrTitle = StrTitle & " Group :" & GroupName
                End If

End If


If CustomerLocation <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrTitle = StrTitle & " ··„Êﬁ⁄ :" & CustomerLocation
                Else
                          StrTitle = StrTitle & " Location :" & CustomerLocation
                End If

End If



    StrFileName = App.path & "\Reports\" & Reportname

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)

        If BolShowNotes = False Then
            xReport.Database.SetDataSource rs
        Else
        
            xReport.Database.Tables(1).SetDataSource rs
            xReport.Database.Tables(2).SetDataSource RsNotes
        End If

        '---------------------------------------------------
        If Reportname = "ReportSalingTimeNotTax.rpt" Then
            xReport.OpenSubreport("SubReport1").Database.SetDataSource rs
            xReport.OpenSubreport("SubRptCustomerTrans.rpt").Database.SetDataSource rs
        End If
    
        Set cOptions = New ClsCompanyInfo
  If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
   Else
   xReport.ParameterFields(1).AddCurrentValue cOptions.EngCompanyName
   End If
  If newReport = 8 Then
    xReport.ParameterFields(6).AddCurrentValue CStr(SystemOptions.Count_ACCOUNT_digit)
  End If
  
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
    
        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If
    
        If Reportname = "ReportSalingTimeNotTax.rpt" Then

            For i = 1 To xReport.ParameterFields.count
                xReport.OpenSubreport("SubReport1").ParameterFields(i).AddCurrentValue xReport.ParameterFields(i).GetNthCurrentValue(1)
            Next i

            For i = 1 To xReport.ParameterFields.count
                xReport.OpenSubreport("SubRptCustomerTrans.rpt").ParameterFields(i).AddCurrentValue xReport.ParameterFields(i).GetNthCurrentValue(1)
            Next i

        End If

        Screen.MousePointer = vbDefault
    End If

    xReport.reporttitle = StrTitle

   

    If newReport = 18 Or newReport = 181 Then
        'x = MsgBox("Â·  —Ìœ ⁄—÷ „ Ê”ÿ «· ﬂ·›…  ‰⁄„ «„ ·« ", vbCritical + vbYesNo)
        'If x = vbYes Then
        xReport.ParameterFields(6).AddCurrentValue ViewCost
        'Else
        'xReport.ParameterFields(6).AddCurrentValue "0"
        'End If
    End If



    If newReport = 600 Then
      '    xReport.ParameterFields(6).AddCurrentValue ViewCost
          xReport.ParameterFields(7).AddCurrentValue ShowZero
    End If
    
    
    If newReport = 1001 Or newReport = 1002 Then
        If ReturnPurcahse <> 0 Then
            xReport.ParameterFields(7).AddCurrentValue "«· ﬁ—Ì— ·« ÌÕ ÊÌ ”‰œ«  ’—› „—œÊœ«  «·„‘ —Ì« "
        End If

        xReport.ParameterFields(7).AddCurrentValue ""

    End If
    
    If mTypeDebitCredit = 1 Then
        xReport.RecordSelectionFormula = "{@netBalanceCredit} = 0"
    ElseIf mTypeDebitCredit = 2 Then
        xReport.RecordSelectionFormula = "{@NetBalanceDebit} = 0"
    End If
    

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , StrFileName, , StrSQL
    Set CViewer = Nothing
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Msg = Msg + CHR(13) & "Err.Description:" & Err.description
    Msg = Msg + CHR(13) & "Err.Source:" & Err.Source
    Msg = Msg + CHR(13) & "Err.Number:" & Err.Number
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function














' ﬁ—Ì— «·„»Ì⁄«  ·⁄„Ì· „⁄Ì‰
Public Function ShowSallingClients(StrSQL As String, _
                                   ClientName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim ShowTax  As Boolean
    Dim Reportname As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „»Ì⁄«  ·Â–« «·⁄„Ì·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    'Œ«’ »÷—Ì»…«·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)

    If ShowTax = True Then
        Reportname = "ReportSallingClient.rpt"
    Else
        Reportname = "ReportSallingClientNotTax.rpt"
    End If

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„»Ì⁄«  «·ÌÊ„Ì…
Public Function SallingDailyReport(StrSQL As String, _
                                   Optional PrintTo As PrintTarget = PrintTarget.WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì „»Ì⁄«  "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\SallingDailyReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\SallingDailyReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\SallingDailyReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, PrintTo
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„»Ì⁄«  ·’‰› „⁄Ì‰
Public Function ShowSallingItems(StrSQL As String, _
                                 Item As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „»Ì⁄«  ·Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportItems.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportItems.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If Item = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Item
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„»Ì⁄«  ·⁄„Ì· „⁄Ì‰ ›Ì › —… „⁄Ì‰…
Public Function ShowSallingClient_Time(StrSQL As String, _
                                       ClientName As String, _
                                       Optional DateFrom As Variant = Null, _
                                       Optional DateTo As Variant = Null)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim ShowTax As Boolean
    Dim Reportname As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „»Ì⁄«  ·Â–« «·⁄„Ì· ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    'Œ«’ »÷—Ì»…«·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)

    If ShowTax = True Then
        Reportname = "ReportClint_Time.rpt"
    Else
        Reportname = "ReportClint_TimeNotTax.rpt"
    End If

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„‘ —Ì«  ›Ì › —… „⁄Ì‰…
Public Function ShowBuyTime(StrSQL As String, _
                            DateFrom As Variant, _
                            DateTo As Variant)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer

    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „‘ —Ì«  ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportBuyTime.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportBuyTime.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportBuyTime.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If IsNull(DateFrom) Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue DisplayDate(CDate(DateFrom))
        End If

        If IsNull(DateTo) Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue DisplayDate(CDate(DateTo))
        End If

        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„‘ —Ì«  ·⁄„Ì· „⁄Ì‰
Public Function ShowBuyClients(StrSQL As String, _
                               ClientName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „‘ —Ì«  „‰ Â–« «·⁄„Ì·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportBuyClient.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportBuyClient.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportBuyClient.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„‘ —Ì«  „‰ ⁄„Ì· „⁄Ì‰ ›Ì › —… „⁄Ì‰…
Public Function ShowBuyClient_Time(StrSQL As String, _
                                   ClientName As String, _
                                   DateFrom As Date, _
                                   DateTo As Date)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì… ‘—«¡ „‰ Â–« «·⁄„Ì· ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportBuyClient_Time.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportBuyClient_Time.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportBuyClient_Time.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„‘ —Ì«  «·ÌÊ„Ì…
Public Function DailyPurchase(StrSQL As String, _
                              Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì ⁄„·Ì«  ‘—«¡"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\PurchaseDailyReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\PurchaseDailyReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\PurchaseDailyReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— „— Ã⁄ «·„‘ —Ì«  «·ÌÊ„Ì…
Public Function DailyReturn(StrSQL As String, _
                            Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì… „— Ã⁄ „‘ —Ì«  «·ÌÊ„ "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DailyReturn.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DailyReturn.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DailyReturn.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— „— Ã⁄ «·„»Ì⁄«  «·ÌÊ„Ì…
Public Function DailyReturnSalling(StrSQL As String, _
                                   Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì… „— Ã⁄ „»Ì⁄«  «·ÌÊ„ "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DailyReturnSalling.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DailyReturnSalling.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DailyReturnSalling.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·„‘ —Ì«  ·’‰› „⁄Ì‰
Public Function ShowBuyingItems(StrSQL As String, _
                                Item As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „‘ —Ì«  ·Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportBuyItems.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportBuyItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportBuyItems.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If Item = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Item
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'÷—«∆» «·„»Ì⁄« 
Public Function SallingTax(StrSQL As String, _
                           Optional ClientName As String, _
                           Optional DateFrom As String, _
                           Optional DateTo As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„ Ì „ Õ”«» ÷—Ì»… „»Ì⁄« " & CHR(13)
        Msg = Msg + "⁄·Ì √Ì ›« Ê—… »Ì⁄"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportTax.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportTax.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportTax.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'÷—«∆» «·„»Ì⁄«  ·› —… „⁄Ì‰…
Public Function SallingTimeTax(StrSQL As String, _
                               Optional DateFrom As String, _
                               Optional DateTo As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„ Ì „ Õ”«» ÷—Ì»… „»Ì⁄« " & CHR(13)
        Msg = Msg + "⁄·Ì √Ì ›« Ê—… »Ì⁄"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportTimeTax.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportTimeTax.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportTimeTax.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'÷—«∆» «·„»Ì⁄«  ·⁄„Ì· „⁄Ì‰
Public Function ClientTax(StrSQL As String, _
                          Optional ClientName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„ Ì „ Õ”«» ÷—Ì»… „»Ì⁄« " & CHR(13)
        Msg = Msg + "⁄·Ì √Ì ›« Ê—… »Ì⁄"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportClientTax.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportClientTax.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportClientTax.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·’Ì«‰…
Public Function MaintenceReport(StrSQL As String, _
                                Optional ClientName As String, _
                                Optional DateFrom As String, _
                                Optional DateTo As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì… ’Ì«‰… ·Â–« «·⁄„Ì· ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportMaintence.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportMaintence.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportMaintence.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·’Ì«‰… «·ÌÊ„Ì…
Public Function MaintenceDailyReport(StrSQL As String, _
                                     Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "··«  ÊÃœ √Ì ⁄„·Ì«  ’Ì«‰… "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\RrportDailyMaintence.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\RrportDailyMaintence.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\RrportDailyMaintence.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·’Ì«‰…
Public Function MaintenceReportClient(StrSQL As String, _
                                      Optional ClientName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì…’Ì«‰… ·Â–« «·⁄„Ì· ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportMaintenceClient.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportMaintenceClient.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportMaintenceClient.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function BalanceSheetAtt(StrSQL As String, _
                                reportno As Integer, _
                                chequeBoxname As String, _
                                titls As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim filename As String

    If reportno = 0 Then
        filename = App.path & "\Reports\BalanceSheetAtt\Report1.rpt"
    ElseIf reportno = 1 Then
        filename = App.path & "\Reports\BalanceSheetAtt\Report2.rpt"
    End If

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
 
        xReport.reporttitle = xReport.reporttitle & CHR(13) & titls
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , filename
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function ChequeReports(StrSQL As String, _
                              reportno As Integer, _
                              chequeBoxname As String, _
                              titls As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ChequeReports\Report1.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ChequeReports\Report1.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ChequeReports\Report1.rpt")
        xReport.Database.SetDataSource rs

        If reportno = 0 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«   «·„ÊÃÊœ… ›Ì Õ«›Ÿ… «·‘Ìﬂ«  " & chequeBoxname
        ElseIf reportno = 1 Then
            xReport.reporttitle = "«·‘Ìﬂ«   «·„ÊÃÊœ… ›Ì Õ«›Ÿ… «·‘Ìﬂ«  Ê „ «Ìœ«⁄Â« " & chequeBoxname
        ElseIf reportno = 2 Then
            xReport.reporttitle = "«·‘Ìﬂ«   «·„ÊÃÊœ… ›Ì Õ«›Ÿ… «·‘Ìﬂ«   Ê·„ Ì „  «Ìœ«⁄Â« " & chequeBoxname
        End If

        xReport.reporttitle = xReport.reporttitle & CHR(13) & titls
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\ChequeReports\Report1.rpt"
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function ChequeReports2(StrSQL As String, _
                               reportno As Integer, _
                               chequeBoxname As String, _
                               titls As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ChequeReports\Report3.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ChequeReports\Report1.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ChequeReports\Report3.rpt")
        xReport.Database.SetDataSource rs

        If reportno = 0 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«     «·„ƒÃ·… «· Ì ·„  ”œœ  " & chequeBoxname
 
        End If

        xReport.reporttitle = xReport.reporttitle & CHR(13) & titls
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\ChequeReports\Report3.rpt"
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function ChequeReports1(StrSQL As String, _
                               reportno As Integer, _
                               chequeBoxname As String, _
                               titls As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ChequeReports\Report2.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ChequeReports\Report1.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ChequeReports\Report2.rpt")
        xReport.Database.SetDataSource rs

        If reportno = 0 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«    «· Ì  „  Õ’Ì·Â« " & chequeBoxname
        ElseIf reportno = 1 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«    «· Ì  „ ”œ«œÂ«  " & chequeBoxname
        ElseIf reportno = 2 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«    «· Ì  «·„— œ… ⁄·Ï «·‘—ﬂ… " & chequeBoxname
        ElseIf reportno = 2 Then
            xReport.reporttitle = "ﬂ· «·‘Ìﬂ«    «· Ì  «·„— œ…   ··‘—ﬂ… " & chequeBoxname
        
        End If

        xReport.reporttitle = xReport.reporttitle & CHR(13) & titls
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , App.path & "\Reports\ChequeReports\Report2.rpt"
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'«·√’‰«› «· Ì  „ «” »œ«·Â«
Public Function MaintenceReplacement(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReplacementReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReplacementReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReplacementReport.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName = "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue ClientName
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·√’‰«›
Public Function ItemsReport(StrSQL As String, Optional ByVal mReportName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If
Dim filename As String

If mReportName = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    filename = App.path & "\Reports\ReportItemsData.rpt"
     Else
     filename = App.path & "\Reports\ReportItemsDataE.rpt"
     End If
Else
    filename = App.path & "\Reports\REPORTS NEW\" & Trim(mReportName)
End If
    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
        xReport.reporttitle = "»Ì«‰«  «·√’‰«›"
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , filename
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

 
 ' ﬁ«—Ì— «·√’‰«›
Public Function ItemsReport3(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If
Dim filename As String
If SystemOptions.UserInterface = ArabicInterface Then
        filename = App.path & "\Reports\ReportItemsData3.rpt"
Else
filename = App.path & "\Reports\ReportItemsData3E.rpt"
End If

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        
        xReport.Database.SetDataSource rs
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.reporttitle = "»Ì«‰«  «·√’‰«›"
     Else
     xReport.reporttitle = "Items Data "
     End If
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , filename, , StrSQL
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·√’‰«›
Public Function ItemsReport2(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If
Dim Fiename As String

If SystemOptions.UserInterface = ArabicInterface Then
Fiename = App.path & "\Reports\ReportGrop.rpt"
Else
Fiename = App.path & "\Reports\ReportGropE.rpt"
End If

    If Dir(Fiename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + Fiename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Fiename)
        xReport.Database.SetDataSource rs
        xReport.reporttitle = "»Ì«‰«  «·√’‰«›"
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Fiename
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·⁄„·Ì«  «· Ì  „  ⁄·Ï «·√’‰«›
Public Function TransReport(StrSQL As String, _
                            Optional BolTotal As Boolean = False, _
                            Optional StrReportCaption As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì ⁄„·Ì«   Œ’ Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If
Dim filename As String
If SystemOptions.UserInterface = ArabicInterface Then
        filename = App.path & "\Reports\ReportItemsTrans.rpt"
Else
filename = App.path & "\Reports\ReportItemsTransE.rpt"
End If
    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.reporttitle = StrReportCaption
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget, xReport.reporttitle, , , , filename
    Set cReportViewer = Nothing
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)

    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·Ã—œ «· ›’Ì·Ì
Public Function CompletlyGard(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportGardComp.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportGardComp.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportGardComp.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·„’—Ê›« 
Public Function ExpensesReports(StrSQL As String, _
                                Optional PrintTo As PrintTarget = WindowTarget, _
                                Optional StrReportDes As String = "", _
                                Optional BolShowDes As Boolean = False, _
                                Optional IntReportStyle As Integer = 0, _
                                Optional ShowChart As Boolean = False, _
                                Optional ChartPostion As Integer = 0)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim cReportViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo

   ' On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If IntReportStyle = 0 Then
        StrFileName = App.path & "\Reports\ReportExpenses.rpt"
    ElseIf IntReportStyle = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\ReportExpensesStyle1.rpt"
     Else
     StrFileName = App.path & "\Reports\ReportExpensesStyle1e.rpt"
     End If
    ElseIf IntReportStyle = 2 Then
    ElseIf IntReportStyle = 3 Then
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportExpenses.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    xReport.reporttitle = " ﬁ—Ì— «·„’—Ê›« "
    Set cCompanyInfo = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue StrReportDes
   ' xReport.ParameterFields(6).AddCurrentValue True

    If BolShowDes = True Then
        xReport.ParameterFields(5).AddCurrentValue 1
    Else
        xReport.ParameterFields(5).AddCurrentValue 0
    End If

    If ShowChart = False Then
        xReport.ParameterFields(6).AddCurrentValue False
        xReport.ParameterFields(7).AddCurrentValue False
    ElseIf ChartPostion = 0 Then
        xReport.ParameterFields(6).AddCurrentValue True
        xReport.ParameterFields(7).AddCurrentValue False
    ElseIf ChartPostion = 1 Then
        xReport.ParameterFields(6).AddCurrentValue False
        xReport.ParameterFields(7).AddCurrentValue True
    End If

    Screen.MousePointer = vbDefault

    If PrintTo = WindowTarget Then
        Set cReportViewer = New ClsReportViewer
        cReportViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Msg = Msg + "" & Err.description & CHR(13)
    Msg = Msg + "" & Err.Number & CHR(13)
    Msg = Msg + "" & Err.Source

    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— „·Œ’ «·⁄„·Ì«  «·ÌÊ„Ì…
Public Function DailSummery(Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If Dir(App.path & "\Reports\ReportSummery.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportSummery.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportSummery.rpt")
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue val(FrmDailtyReport.LblVal(0).Caption)
        xReport.ParameterFields(5).AddCurrentValue val(FrmDailtyReport.lblDiscount(0).Caption)
        xReport.ParameterFields(6).AddCurrentValue val(FrmDailtyReport.LblTotal(0).Caption)
        xReport.ParameterFields(7).AddCurrentValue val(FrmDailtyReport.LblVal(1).Caption)
        xReport.ParameterFields(8).AddCurrentValue val(FrmDailtyReport.lblDiscount(1).Caption)
        xReport.ParameterFields(9).AddCurrentValue val(FrmDailtyReport.LblTotal(1).Caption)
        xReport.ParameterFields(15).AddCurrentValue val(FrmDailtyReport.LblVal(8).Caption)
        xReport.ParameterFields(10).AddCurrentValue val(FrmDailtyReport.LblVal(2).Caption)
        xReport.ParameterFields(11).AddCurrentValue val(FrmDailtyReport.LblVal(3).Caption)
        xReport.ParameterFields(12).AddCurrentValue val(FrmDailtyReport.LblVal(4).Caption)
        xReport.ParameterFields(13).AddCurrentValue val(FrmDailtyReport.LblVal(6).Caption)
        xReport.ParameterFields(14).AddCurrentValue val(FrmDailtyReport.LblVal(5).Caption)
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·„œ›Ê⁄« 
Public Function PaymentsReports(StrSQL As String, _
                                Optional PrintTo As PrintTarget = WindowTarget, _
                                Optional StrReportDes As String = "", _
                                Optional BolShowDes As Boolean = False)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\PaymentsReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\PaymentsReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\Reports\PaymentsReport.rpt")
        Else
        Set xReport = xApp.OpenReport(App.path & "\Reports\PaymentsReporte.rpt")
        End If
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue Trim(StrReportDes)

        If BolShowDes = True Then
            xReport.ParameterFields(5).AddCurrentValue 1
        Else
            xReport.ParameterFields(5).AddCurrentValue 0
        End If

        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        If SystemOptions.UserInterface = ArabicInterface Then
                FrmReport.txtPath = (App.path & "\Reports\PaymentsReport.rpt")
         Else
                FrmReport.txtPath = (App.path & "\Reports\PaymentsReporte.rpt")
         End If
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·„ﬁ»Ê÷« 
Public Function CashingReports(StrSQL As String, _
                               Optional PrintTo As PrintTarget = WindowTarget, _
                               Optional StrReportDes As String = "", _
                               Optional BolShowDes As Boolean = False)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String
    Dim CViewer As ClsReportViewer

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.path & "\Reports\CahingReport.rpt"
Else
StrFileName = App.path & "\Reports\CahingReporte.rpt"
End If
    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\CahingReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.reporttitle = StrReportDes ' " ﬁ«—Ì— «·„ﬁ»Ê÷« "
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue StrReportDes

        If BolShowDes = True Then
            xReport.ParameterFields(5).AddCurrentValue 1
        Else
            xReport.ParameterFields(5).AddCurrentValue 0
        End If

        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , StrFileName, , StrSQL
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·Ã—œ «·«Ã„«·Ì
Public Function ShortGard(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim StrReportTitle As String
    Dim cReportViewer As ClsReportViewer
    Dim cOptions As ClsCompanyInfo
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    StrFileName = App.path & "\Reports\ReportGardShort.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    StrReportTitle = " ﬁ—Ì— »Ã—œ «·„Œ“Ê‰ Õ Ï " & ""
    xReport.reporttitle = StrReportTitle
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function MaintencePerventReport(DateFrom As Date, _
                                       DateTo As Date)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim sql As String

    On Error GoTo ErrTrap
    sql = "SELECT     dbo.TBLRegularMaint.DateOfRegularMaint, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.Transactions.Transaction_Date, "
    sql = sql & "  dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TBLRegularMaint.GranteeType, dbo.TBLRegularMaint.GranteeStartDate,"
    sql = sql & " dbo.TBLRegularMaint.GranteeEndDate, dbo.TBLRegularMaint.Done, dbo.TBLRegularMaint.DoneDate, dbo.TBLRegularMaint.ItemSerial,"
    sql = sql & " dbo.Transactions.NoteSerial1"
    sql = sql & " FROM         dbo.TBLRegularMaint INNER JOIN"
    sql = sql & " dbo.TblItems ON dbo.TBLRegularMaint.itemid = dbo.TblItems.ItemID INNER JOIN"
    sql = sql & "  dbo.Transactions ON dbo.TBLRegularMaint.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    sql = sql & "     dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
    sql = sql & "  WHERE     (dbo.TBLRegularMaint.DateOfRegularMaint >=" & SQLDate(DateFrom, True) & " )"
    sql = sql & "  and      (dbo.TBLRegularMaint.DateOfRegularMaint <=" & SQLDate(DateTo, True) & " )"
 
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«  ÊÃœ Œÿ… ’Ì«‰… ›Ì Â–… «·› —… "
        Else
            Msg = "There are no preventive maintenance plan in this Period  "
        End If
    
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    filename = "\Reports\ReportPerventMaintence.rpt"

    If Dir(App.path & filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & filename)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    FrmReport.txtPath = (App.path & filename)
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function MaintenceReportTime(StrSQL As String, _
                                    Optional DateFrom As String, _
                                    Optional DateTo As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì…’Ì«‰… ·Â–« «·⁄„Ì· ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportMaintenceTime.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportMaintenceTime.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportMaintenceTime.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If DateFrom = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(6).AddCurrentValue ""
        Else
            xReport.ParameterFields(6).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— ·«Â·«ﬂ«  «· Ì  „  ›Ì › —… „⁄Ì‰…
Public Function DestructionReportTime(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer

    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionTime.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionTime.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionTime.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function DestructionReportTimeNew(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer

    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionItemsNew.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionItemsNew.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionItemsNew.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·«Â·«ﬂ«  ·’‰› „⁄Ì‰
Public Function DestructionItems(StrSQL As String, _
                                 IntType As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ »Ì«‰«  ··⁄—÷...!!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionItems.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionItems.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If IntType = 8 Then
            xReport.reporttitle = " ﬁ—Ì— ’—› »÷«⁄… -  ·›Ì«   "
        ElseIf IntType = 17 Then
            xReport.reporttitle = " ﬁ—Ì— ’—› »÷«⁄…- „”ÕÊ»«  ‘Œ’Ì…"
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'  ﬁ«—Ì— «·√’‰«› «· Ì Ê’·  Õœ «·ÿ·»
Public Function RequestItems(StrSQL As String, _
                             Optional BolTreeStyle As Boolean = False)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim CViewer As ClsReportViewer
    Dim StrFileName As String
    Dim cOptions As ClsCompanyInfo
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    If BolTreeStyle = False Then
        
      If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\RequestItems.rpt"
      Else
      StrFileName = App.path & "\Reports\RequestItemsE.rpt"
      End If
    Else
        StrFileName = App.path & "\Reports\RequestItemsTreeStyle.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                      Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
                Else
                Msg = "No Data"
                End If
  
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    If SystemOptions.UserInterface = ArabicInterface Then
    xReport.reporttitle = " ﬁ—Ì— »«·√’‰«› «· Ï Ê’·  Õœ «·ÿ·»"
    Else
    xReport.reporttitle = "On Demand Items"
    End If
    Set cOptions = New ClsCompanyInfo
    xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
    xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
    xReport.ParameterFields(3).AddCurrentValue user_name
    Screen.MousePointer = vbDefault
   
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , StrFileName
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'ﬂ‘› Õ”«» ⁄„Ì· («·ﬁÌ„ «·„ƒÃ·…⁄·Ï «·⁄„Ì·)
Public Function ClientDelayValue(StrSQL As String, _
                                 CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ClientDelayVal.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ClientDelayVal.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ClientDelayVal.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'ﬂ‘› Õ”«» ⁄„Ì· («·ﬁÌ„ «·„ƒÃ·… ··⁄„Ì·)
Public Function ClientDelayValue_To(StrSQL As String, _
                                    CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ClientDelayVal_to.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ClientDelayVal_to.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ClientDelayVal_to.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

''ﬂ‘› Õ”«» ⁄„Ì· («·„ﬁ»Ê÷« )
Public Function ClientCashing(StrSQL As String, _
                              CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ClientCaching.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ClientCaching.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ClientCaching.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

''ﬂ‘› Õ”«» ⁄„Ì· («·„œ›Ê⁄« )
Public Function ClientPayments(StrSQL As String, _
                               CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\Client_PayTo.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\Client_PayTo.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Client_PayTo.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'ﬂ‘› Õ”«» „Ê—œ («·ﬁÌ„ «·„ƒÃ·…⁄·Ï «·„Ê—œ)
Public Function SuplierDelayValue(StrSQL As String, _
                                  CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\Sup_DelayVal.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\Sup_DelayVal.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Sup_DelayVal.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'ﬂ‘› Õ”«» „Ê—œ («·ﬁÌ„ «·„ƒÃ·… ··„Ê—œ)
Public Function Sup_DelayValue_To(StrSQL As String, _
                                  CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\SupDelayVal_to.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\SupDelayVal_to.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\SupDelayVal_to.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

''ﬂ‘› Õ”«» „Ê—œ («·„ﬁ»Ê÷« )
Public Function Suplier_Cashing(StrSQL As String, _
                                CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\Sup_Cashing.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\Sup_Cashing.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Sup_Cashing.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

''ﬂ‘› Õ”«» „Ê—œ («·„œ›Ê⁄« )
Public Function Sup_Payments(StrSQL As String, _
                             CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\Sup_PayTo.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\Sup_PayTo.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Sup_PayTo.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

''ﬂ‘› Õ”«» „Ê—œ
Public Function Suplier_Account(StrSQL As String, _
                                CusName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\Sup_Account.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\Sup_Account.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\Sup_Account.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ«—Ì— «·„œ›Ê⁄«  «·ÌÊ„Ì…
Public Function PaymentsDailyReport(StrSQL As String, _
                                    Optional PrintTo As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrFileName As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷ "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\PaymentsReport.rpt"
    Else
        StrFileName = App.path & "\Reports\PaymentsReport_Eng.rpt"
    End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.txtPath = StrFileName
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'«·ﬁÌ„ «·¬Ã·… ⁄·Ï «·⁄„·«¡ Ê«·„Ê—œÌ‰
Public Function DelayValue(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ClientsDelayValues.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ClientsDelayValues.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ClientsDelayValues.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    If PrintTo = WindowTarget Then
        Set FrmReport = New FrmReportViewer
        FrmReport.CRViewer.ReportSource = xReport
        FrmReport.CRViewer.ViewReport
        FrmReport.show
        Screen.MousePointer = vbDefault
        SendKeys "{RIGHT}"
    ElseIf PrintTo = PrinterTarget Then
        xReport.PrintOut False
    End If

    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'«·ﬁÌ„ «·¬Ã·… ⁄·Ï «·‘—ﬂ…
Public Function DelayValueOnCompany(StrSQL As String, _
                                    Optional m_Target As PrintTarget = WindowTarget)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim StrHeader As String
    Dim cReportViewer As ClsReportViewer

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\CompanyDelayVal.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\CompanyDelayVal.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\CompanyDelayVal.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue CusName
        Screen.MousePointer = vbDefault
    End If

    StrHeader = "«·ﬁÌ„ Ê«·√Ê—«ﬁ «·„«·Ì… «·√Ã·… ⁄·Ï «·‘—ﬂ… "
    xReport.reporttitle = StrHeader
    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Target, " ﬁ«—Ì— «·„⁄«„·«  «·„«·Ì…"
    Set cReportViewer = Nothing
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function
Public Function NotorReceiptedPart(StrSQL As String, _
                                 Optional ClientName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "··«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\REPORTS\REPORTS NEW\" & "QesTabel.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\QesNotReceipted.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\REPORTS\REPORTS NEW\" & "QesTabel.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
            xReport.ParameterFields(5).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue "1"
            xReport.ParameterFields(5).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.txtPath = App.path & "\REPORTS\REPORTS NEW\" & "QesTabel.rpt"
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function
' ﬁ—Ì— «·√ﬁ”«ÿ «· Ì  „  Õ’Ì·Â«
Public Function ReceiptedPart(StrSQL As String, _
                              IntReceiptType As Integer, _
                              Optional ClientName As String = "", _
                              Optional m_Target As PrintTarget = WindowTarget)
    'IntReceiptType=0 ' Õ’Ì· √ﬁ”«ÿ
    'IntReceiptType=1 '”œ«œ √ﬁ”«ÿ
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim cReportViewer As ClsReportViewer

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\QestHadReceipted.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\QestHadReceipted.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\QestHadReceipted.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
            xReport.ParameterFields(5).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue "1"
            xReport.ParameterFields(5).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    If IntReceiptType = 0 Then
        xReport.reporttitle = " ﬁ—Ì— »«·√ﬁ”«ÿ «·„Õ’·…"
    ElseIf IntReceiptType = 1 Then
        xReport.reporttitle = " ﬁ—Ì— »«·√ﬁ”«ÿ «·„”œœ…"
    End If

    Set cReportViewer = New ClsReportViewer
    cReportViewer.FireReport xReport, m_Target, " ﬁ—Ì— «·√ﬁ”«ÿ ", , , , App.path & "\Reports\QestHadReceipted.rpt"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— «·√ﬁ”«ÿ «· Ì ·„  Ì „  Õ’Ì·Â«
Public Function NotReceiptedPart(StrSQL As String, _
                                 Optional ClientName As String = "")
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "··«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\QesNotReceipted.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\QesNotReceipted.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\QesNotReceipted.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
            xReport.ParameterFields(5).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue "1"
            xReport.ParameterFields(5).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.txtPath = App.path & "\Reports\QesNotReceipted.rpt"
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'      ÿ·»«   „ «⁄ „«œÂ« ··„»Ì⁄« 
Public Function AccreditOrders(StrSQL As String, _
                               Optional ClientName As String = "", _
                               Optional mytitle As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap
    Dim filename As String

    If StrSQL = Null Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì        ÿ·»«   „ «⁄ „«œÂ« ··„»Ì⁄« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    filename = App.path & "\Reports\New Reports\AccreditOrders.rpt"

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = mytitle
 
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    FrmReport.txtPath = filename

    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'√ﬁ”«ÿ Õ«‰ Êﬁ  ”œ«œÂ«
Public Function QestMustPayed(StrSQL As String, _
                              Optional ClientName As String = "", _
                              Optional mytitle As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap
    Dim filename As String

    If StrSQL = Null Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ √Ì √ﬁ”«ÿ Õ«‰ Êﬁ  ”œ«œÂ«"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    filename = App.path & "\Reports\QestMustPayed.rpt"

    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + filename
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\QestMustPayed.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.reporttitle = mytitle

        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
            xReport.ParameterFields(5).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue "1"
            xReport.ParameterFields(5).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    FrmReport.txtPath = filename

    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— ≈Ã„«·Ì «·√ﬁ”«ÿ «· Ì  „  Õ’Ì·Â«
Public Function ReceiptedPart_Total(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„ Ì „ ”œ«œ √Ì √ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\QestHadReceiptedTotal.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\QestHadReceiptedTotal.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\QestHadReceiptedTotal.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.txtPath = App.path & "\Reports\QestHadReceiptedTotal.rpt"
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

' ﬁ—Ì— ≈Ã„«·Ì «·√ﬁ”«ÿ «· Ì ·„ Ì „  Õ’Ì·Â«
Public Function NotReceiptedPart_Total(StrSQL As String)

    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "··«  ÊÃœ √Ì »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\QesNotReceiptedTotal.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\QesNotReceiptedTotal.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\QesNotReceiptedTotal.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name

        If ClientName <> "" Then
            xReport.ParameterFields(4).AddCurrentValue ""
            xReport.ParameterFields(5).AddCurrentValue ClientName
        Else
            xReport.ParameterFields(4).AddCurrentValue "1"
            xReport.ParameterFields(5).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.txtPath = App.path & "\Reports\QesNotReceiptedTotal.rpt"
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Function

' ﬁ—Ì— „— Ã⁄ «·„»Ì⁄« 
 

Public Function ReturnSallingData(SallingID As Long, Optional LblTotal As String, Optional Custcode As String, Optional BranchVATNO As String, Optional BranchID As Integer, Optional repType As Integer, Optional ByVal X As Integer = 0)
    
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrSQL As String
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap
    On Error Resume Next
    If SallingID = Null Then
        Exit Function
    End If

    'StrSQL = "select * From QryBuyReport where Transaction_ID=" & SallingID

    StrSQL = " SELECT   dbo.Transactions.ReturnSerial ,  dbo.Transactions.TransactionComment , dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, "
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                  dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName,"
    StrSQL = StrSQL & "                  dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
    StrSQL = StrSQL & "                  dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
    StrSQL = StrSQL & "                  dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblCustemers.Fullcode,"
    StrSQL = StrSQL & "                  dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Nots, dbo.Transactions.Nots2,"
    StrSQL = StrSQL & "                  dbo.Transactions.NoteSerial, dbo.Transactions.Product_Issue_voucher_Serial, dbo.Transactions.WorkOrderNO, Transactions_1.NoteSerial1 AS IssueVoucherNo,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Width, dbo.Transactions.VAT, dbo.Transactions.VATNO, dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo,"
    StrSQL = StrSQL & "                  dbo.Transactions.ResonVAT, dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE,"
    StrSQL = StrSQL & "                  dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & "    FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions Transactions_1 ON dbo.Transactions.Nots = Transactions_1.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID

    StrSQL = StrSQL & " ORDER BY ID "

'new
StrSQL = "SELECT       dbo.Transactions.TransactionComment, dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, "
StrSQL = StrSQL & "                      dbo.Transaction_Details.LotNO, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, TblUnites_1.UnitName,"
StrSQL = StrSQL & "                      dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
StrSQL = StrSQL & "                      dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
StrSQL = StrSQL & "                      dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblCustemers.Fullcode,"
StrSQL = StrSQL & "                      dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Nots, dbo.Transactions.Nots2,"
StrSQL = StrSQL & "                      dbo.Transactions.NoteSerial, dbo.Transactions.Product_Issue_voucher_Serial, dbo.Transactions.WorkOrderNO, Transactions_1.NoteSerial1 AS IssueVoucherNo,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Width, dbo.Transactions.VAT, dbo.Transactions.VATNO, dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo,"
StrSQL = StrSQL & "                      dbo.Transactions.ResonVAT, dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE,"
StrSQL = StrSQL & "                      TblUnites_1.UnitNamee , dbo.TblEmployee.emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions Transactions_1 ON dbo.Transactions.Nots = Transactions_1.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites TblUnites_2 ON dbo.Transaction_Details.UnitId = TblUnites_2.UnitID"

StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID
StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID"

'new 22 07 2019
StrSQL = "SELECT      dbo.Transactions.ReturnSerial ,  dbo.Transactions.TransactionComment, dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, "
StrSQL = StrSQL & "                        dbo.Transaction_Details.LotNO, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID, dbo.Transactions.CusID, dbo.Transactions.StoreID,"
StrSQL = StrSQL & "                        dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.UserID, dbo.TblUsers.UserName, dbo.Transactions.Emp_ID,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.Price,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.discountvalue, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemPhoto, dbo.TblItems.ItemNamee, TblUnites_1.UnitName,"
StrSQL = StrSQL & "                        dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblItemsclasses.SizeName AS ClassName, dbo.Transactions.CashCustomerName,"
StrSQL = StrSQL & "                        dbo.Transactions.NoteSerial1, dbo.Transaction_Details.TotalDiscountPerLine, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
StrSQL = StrSQL & "                        dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType, dbo.Transactions.ManualNo1, dbo.Transactions.PaymentType, dbo.TblCustemers.Fullcode,"
StrSQL = StrSQL & "                        dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.Nots, dbo.Transactions.Nots2,"
StrSQL = StrSQL & "                        dbo.Transactions.NoteSerial, dbo.Transactions.Product_Issue_voucher_Serial, dbo.Transactions.WorkOrderNO, Transactions_1.NoteSerial1 AS IssueVoucherNo,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.ReminRequ, dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, dbo.Transaction_Details.Height,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.Width, dbo.Transactions.VAT, dbo.Transactions.VATNO, dbo.Transaction_Details.Vat AS VatDet, dbo.Transaction_Details.Vatyo,"
StrSQL = StrSQL & "                        dbo.Transactions.ResonVAT, dbo.Transaction_Details.QtyFaqtors, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
StrSQL = StrSQL & "                        dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.MaxUnitID, TblUnites_1.UnitName AS MaxUnitName, TblUnites_1.UnitNamee AS MaxUnitNameE,"
StrSQL = StrSQL & "                        TblUnites_1.UnitNamee , dbo.TblEmployee.emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblUsers ON dbo.Transactions.UserID = dbo.TblUsers.UserID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblUnites TblUnites_1 ON dbo.Transaction_Details.MaxUnitID = TblUnites_1.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.Transactions Transactions_1 ON dbo.Transactions.Nots = Transactions_1.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblUnites TblUnites_2 ON dbo.Transaction_Details.UnitId = TblUnites_2.UnitID"
'Where (dbo.transactions.Transaction_ID = 2340)
'ORDER BY dbo.Transaction_Details.ID
StrSQL = StrSQL & "  where Transactions.Transaction_ID=" & SallingID
StrSQL = StrSQL & " ORDER BY dbo.Transaction_Details.ID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim filename As String
    
    If repType = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'FileName = "ReturnSalling.rpt"
            filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnSalling.rpt"
        Else
            'FileName = "ReturnSallingEng.rpt"
            filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnSallingeng.rpt"
        End If
    ElseIf repType = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
            'FileName = "ReturnSalling.rpt"
            filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnSalling2.rpt"
        Else
            'FileName = "ReturnSallingEng.rpt"
            filename = App.path & "\Special\" & SystemOptions.Reportpath & "\ReturnSalling2E.rpt"
        End If
    End If
    
   Dim report0 As String
Dim report1 As String

                If report1 <> "" Then
                         filename = App.path & "\Special\" & SystemOptions.Reportpath & "\" & report1 & ".rpt"
                End If
                 
                 
    
    If Dir(filename) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = "\Reports\ReturnSalling.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
    Dim cOptions As ClsCompanyInfo
    Set cOptions = New ClsCompanyInfo
        Screen.MousePointer = vbArrowHourglass
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(filename)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(9).AddCurrentValue Custcode
        xReport.ParameterFields(10).AddCurrentValue BranchVATNO
    If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(11).AddCurrentValue cOptions.VATRegNo
    Else
    xReport.ParameterFields(11).AddCurrentValue GetRegVATNo(BranchID)
    End If
        
        xReport.ParameterFields(8).AddCurrentValue WriteNo(CStr(LblTotal), 0, True)
        StrSQL1 = "select * From Notes where Transaction_ID=" & SallingID
        StrSQL1 = StrSQL1 + " and NoteType=0"
        RsTemp.Open StrSQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(RsTemp("Note_Value").value), "", Trim(RsTemp("Note_Value").value))
        Else
            xReport.ParameterFields(4).AddCurrentValue ""
        End If

        Screen.MousePointer = vbDefault
    End If

    Dim StrPrinterName As String


    If (SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3) And X = 0 Then
 
    X = MsgBox("ÿ»«⁄Â „»«‘—Â", vbInformation + vbYesNo)
   
    End If

  
    If (SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3) And X = vbYes Then
        StrPrinterName = DefaultPrinter
        Set cPrinter = New ClsPrinters

        If StrPrinterName <> "" Then
            Set ObjPrinter = cPrinter.GetPrinter(StrPrinterName)
         
        Else
            Set ObjPrinter = Printer
        End If

        If Not ObjPrinter Is Nothing Then
            xReport.SelectPrinter ObjPrinter.drivername, ObjPrinter.DeviceName, ObjPrinter.Port
        End If
                                        
        xReport.PrintOutEx False, IntCopies
     
        Set ObjPrinter = Nothing
        Set cPrinter = Nothing
 
        Exit Function
    End If
   
    '    Screen.MousePointer = vbArrowHourglass
    CreateLogo xReport
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.txtPath = filename
    FrmReport.TXTSTRSQL = StrSQL
    FrmReport.show
        
    Screen.MousePointer = vbDefault
    'SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Else
        Msg = "Sorry... " & CHR(13) & "Can not print this report" & CHR(13)
    
    End If
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function


' ﬁ—Ì— »⁄„·Ì«  «·’Ì«‰… «· Ì  „  ⁄·Ï ﬁÿ⁄… „⁄Ì‰…
Public Function SerialMaintenance(StrSQL As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·„   „ √Ì ⁄„·Ì… ’Ì«‰… ·Â–Â «·ﬁÿ⁄…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\SerialMaintenance.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\SerialMaintenance.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\SerialMaintenance.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

'«· ﬁ—Ì— «·Œ «„Ì
Public Function FinalReport()
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If Dir(App.path & "\Reports\FinalReport.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\FinalReport.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\FinalReport.rpt")
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        '≈Ã„«·Ì «·„»Ì⁄« 
        StrSQL = "SELECT Sum(ReportSallingTime.TotalAfterTax) AS TotalAfterTax FROM ReportSallingTime"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(4).AddCurrentValue val(RsTemp("TotalAfterTax").value)
        End If

        RsTemp.Close
        '≈Ã„«·Ì «·„‘ —Ì« 
        StrSQL = "SELECT Sum(ReportBuyTime_Client.Total) AS Total FROM ReportBuyTime_Client"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(5).AddCurrentValue val(RsTemp("Total").value)
        End If

        RsTemp.Close
        '„— Ã⁄ «·„‘ —Ì« 
        StrSQL = "SELECT Sum(QryReturn.Total) AS Total FROM QryReturn"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(6).AddCurrentValue val(RsTemp("Total").value)
        End If

        RsTemp.Close
        
        '„— Ã⁄ «·„»Ì⁄« 
        StrSQL = "SELECT Sum(QryReturnSalling.Total) AS Total FROM QryReturnSalling "
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(7).AddCurrentValue val(RsTemp("Total").value)
        End If

        RsTemp.Close
        
        '«· ·›Ì« 
        StrSQL = "SELECT Sum(DestructionTime.Summition) AS Summition FROM DestructionTime "
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(8).AddCurrentValue val(RsTemp("Summition").value)
        End If

        RsTemp.Close
        
        '«·„’—Ê›« 
        StrSQL = "SELECT Sum(ExpensesReport.Note_Value) AS Note_Value FROM ExpensesReport "
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(9).AddCurrentValue val(RsTemp("Note_Value").value)
        End If

        RsTemp.Close
        
        '«·„œ›Ê⁄« 
        StrSQL = "SELECT Sum(PaymentsReport.Note_Value) AS Note_Value FROM PaymentsReport"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(10).AddCurrentValue val(RsTemp("Note_Value").value)
        End If

        RsTemp.Close
        
        '«·„ﬁ»Ê÷« 
        StrSQL = "SELECT Sum(CahingReport.Note_Value) AS Note_Value FROM CahingReport"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            xReport.ParameterFields(11).AddCurrentValue val(RsTemp("Note_Value").value)
        End If

        RsTemp.Close
        
        '        xReport.ParameterFields(6).AddCurrentValue 1
        '        xReport.ParameterFields(7).AddCurrentValue 1
        '        xReport.ParameterFields(8).AddCurrentValue 1
        '        xReport.ParameterFields(9).AddCurrentValue 1
        '        xReport.ParameterFields(15).AddCurrentValue 1
        '        xReport.ParameterFields(10).AddCurrentValue 1
        '        xReport.ParameterFields(11).AddCurrentValue 1
        '        xReport.ParameterFields(12).AddCurrentValue 1
        '        xReport.ParameterFields(13).AddCurrentValue 1
        '        xReport.ParameterFields(14).AddCurrentValue 1
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Sub ShowMaintanceStock()
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim StrFileName As String

    On Error GoTo ErrTrap
    StrSQL = "Select * From QryMaintenenceSearchItems Where GoOut=1"
    rs.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    StrFileName = App.path & "\Reports\MaintanceStocke.rpt"

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(StrFileName)
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Sub

Public Function ShowReSales(StrSQL As String, _
                            Optional DateFrom As Variant = Null, _
                            Optional DateTo As Variant = Null, _
                            Optional BolShowNotes As Boolean = False, Optional reportString As String)
    ' ﬁ—Ì— „— Ã⁄ «·„»Ì⁄« 
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Reportname As String
    Dim ShowTax As Boolean
    Dim Msg As String
    Dim RsNotes As ADODB.Recordset
    Dim MySQL As String
    Dim StrTitle As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

If SystemOptions.UserInterface = ArabicInterface Then
    Reportname = "ReSalingTime.rpt"
Else
Reportname = "ReSalingTimeE.rpt"
End If
   
   
   Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        
        Msg = "·«  ÊÃœ „— Ã⁄ „»Ì⁄«  ›Ì Â–Â «·› —…"
Else
Msg = " No Data In this Period"
End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If BolShowNotes = True Then
        StrTitle = " ﬁ—Ì— »„— Ã⁄ «·„»Ì⁄«  «·√Ã·… Ê«· Õ’Ì·«  «·ﬁ«∆„… ⁄·ÌÂ«"
        Set RsNotes = New ADODB.Recordset
        MySQL = "Select * From QryPostedTransNotes Where NoteType=4 OR NoteType=9"
        RsNotes.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
    If SystemOptions.UserInterface = ArabicInterface Then

        StrTitle = " ﬁ—Ì— »„— Ã⁄ «·„»Ì⁄« "
Else
StrTitle = "Sales Return Report"
End If
    
    
    
    End If

If reportString <> "" Then
StrTitle = StrTitle & CHR(13) & " ··„‰œÊ» " & reportString
End If

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)

        If BolShowNotes = False Then
            xReport.Database.SetDataSource rs
        Else
            
            xReport.Database.Tables(1).SetDataSource rs
            xReport.Database.Tables(2).SetDataSource RsNotes
        End If
        
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
   xReport.ParameterFields(6).AddCurrentValue CStr(SystemOptions.Count_ACCOUNT_digit)
        
        
        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    xReport.reporttitle = StrTitle
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    
    FrmReport.show
    FrmReport.txtPath = App.path & "\Reports\" & Reportname
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Function

Public Function ShowRePurchase(StrSQL As String, _
                               Optional DateFrom As Variant = Null, _
                               Optional DateTo As Variant = Null, _
                               Optional BolShowNotes As Boolean = False)
    ' ﬁ—Ì— „— Ã⁄ «·„‘ —Ì« 
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Reportname As String
    Dim ShowTax As Boolean
    Dim Msg As String
    Dim RsNotes As ADODB.Recordset
    Dim MySQL As String
    Dim StrTitle As String

    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

If SystemOptions.UserInterface = ArabicInterface Then
    Reportname = "RePurchaseTime.rpt"
Else
Reportname = "RePurchaseTimeE.rpt"
End If
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«  ÊÃœ „— Ã⁄ „‘ —Ì«  ›Ì Â–Â «·› —…"
            Else
            Msg = "No Data"
            End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If BolShowNotes = True Then
        StrTitle = " ﬁ—Ì— »„— Ã⁄ «·„‘ —Ì«  «·√Ã·… Ê«· Õ’Ì·«  «·ﬁ«∆„… ⁄·ÌÂ«"
        Set RsNotes = New ADODB.Recordset
        MySQL = "Select * From QryPostedTransNotes Where NoteType=4 OR NoteType=9"
        RsNotes.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
                If SystemOptions.UserInterface = ArabicInterface Then

        StrTitle = " ﬁ—Ì— »„— Ã⁄ «·„‘ —Ì« "
        Else
        StrTitle = "Return Purchases"
        End If
    End If

    If Dir(App.path & "\Reports\" & Reportname) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\" & Reportname
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\" & Reportname)

        If BolShowNotes = False Then
            xReport.Database.SetDataSource rs
        Else
            
            xReport.Database.Tables(1).SetDataSource rs
            xReport.Database.Tables(2).SetDataSource RsNotes
        End If
        
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        
        If DateFrom = Null Then
            xReport.ParameterFields(4).AddCurrentValue ""
        Else
            xReport.ParameterFields(4).AddCurrentValue Format(DateFrom, "yyyy/m/d")
        End If

        If DateTo = Null Then
            xReport.ParameterFields(5).AddCurrentValue ""
        Else
            xReport.ParameterFields(5).AddCurrentValue Format(DateTo, "yyyy/m/d")
        End If

        Screen.MousePointer = vbDefault
    End If

    xReport.reporttitle = StrTitle
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
FrmReport.txtPath = App.path & "\Reports\" & Reportname

    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Function

Public Function DestructionTotal(StrSQL As String, _
                                 Optional fromdte As String, _
                                 Optional ToDate As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ »Ì«‰«  ··⁄—÷...!!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionItems_Total.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionItems_Total.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        '        xReport.ParameterFields(5).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").Value), "", StrRS("Company_Comment").Value)
        '        xReport.ParameterFields(3).AddCurrentValue User_Name
        xReport.ParameterFields(1).AddCurrentValue " ﬁ—Ì— ⁄‰ «Ã„«·Ì«  «Ê«„— «·‘€·"
        xReport.ParameterFields(2).AddCurrentValue fromdte
        xReport.ParameterFields(3).AddCurrentValue ToDate
        '        If IntType = 8 Then
        '            xReport.ReportTitle = " ﬁ—Ì— ’—› »÷«⁄…( ·›Ì« )"
        '        ElseIf IntType = 17 Then
        '            xReport.ReportTitle = " ﬁ—Ì— ’—› »÷«⁄…(„”ÕÊ»«  ‘Œ’Ì…)"
        '        End If
        Screen.MousePointer = vbDefault
    End If
    
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function DestructionTotCost(StrSQL As String, _
                                   fromdte As String, _
                                   ToDate As String, _
                                   OrdNUM As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = Null Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ »Ì«‰«  ··⁄—÷...!!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionItems_TotCost.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionItems.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionItems_TotCost.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").Value), "", StrRS("Company_Arabic_Name").Value)
        '        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").Value), "", StrRS("Company_Comment").Value)
        '        xReport.ParameterFields(3).AddCurrentValue User_Name
        xReport.ParameterFields(1).AddCurrentValue " ﬁ—Ì— ⁄‰ «Ã„«·Ì«  «·«’‰«›"
        xReport.ParameterFields(2).AddCurrentValue fromdte
        xReport.ParameterFields(3).AddCurrentValue ToDate
        xReport.ParameterFields(4).AddCurrentValue OrdNUM
        
        '        If IntType = 8 Then
        '            xReport.ReportTitle = " ﬁ—Ì— ’—› »÷«⁄…( ·›Ì« )"
        '        ElseIf IntType = 17 Then
        '            xReport.ReportTitle = " ﬁ—Ì— ’—› »÷«⁄…(„”ÕÊ»«  ‘Œ’Ì…)"
        '        End If
        Screen.MousePointer = vbDefault
    End If
    
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.CRViewer.ViewReport
    FrmReport.show
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function ShowBuyTime_1(StrSQL As String, _
                              DateFrom As Variant, _
                              DateTo As Variant, _
                              ItemName As String)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer

    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·«  ÊÃœ „‘ —Ì«  ›Ì Â–Â «·› —…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\ReportBuyItems_1.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\ReportBuyTime.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\ReportBuyItems_1.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockReadOnly, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue ItemName

        '        If IsNull(DateFrom) Then
        '            xReport.ParameterFields(4).AddCurrentValue ""
        '        Else
        '            xReport.ParameterFields(4).AddCurrentValue DisplayDate(CDate(DateFrom))
        '        End If
        '        If IsNull(DateTo) Then
        '            xReport.ParameterFields(5).AddCurrentValue ""
        '        Else
        '            xReport.ParameterFields(5).AddCurrentValue DisplayDate(CDate(DateTo))
        '        End If
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    SendKeys "{RIGHT}"
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Public Function DestructionItems_Diff(StrSQL As String, _
                                      IntType As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrRS As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer

    Dim Msg As String
    On Error GoTo ErrTrap

    If StrSQL = "" Then
        Exit Function
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    If Dir(App.path & "\Reports\DestructionTime_Diff.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + "\Reports\DestructionTime_Diff.rpt"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(App.path & "\Reports\DestructionTime_Diff.rpt")
        xReport.Database.SetDataSource rs
        StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
        xReport.ParameterFields(3).AddCurrentValue user_name
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

