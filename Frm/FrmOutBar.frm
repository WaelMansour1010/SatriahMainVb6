VERSION 5.00
Begin VB.Form FrmOutBar 
   BorderStyle     =   0  'None
   Caption         =   "‘—Ìÿ «·«Œ ’«—« "
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FrmOutBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo ErrTrap
Dim BGround As ClsBackGroundPic
Set BGround = New ClsBackGroundPic
OutBar.DefaultBkGround = BGround.Picture
Me.OutBar.ChangingSpeed = 0
Me.OutBar.FileDragDrop = True
Exit Sub
ErrTrap:
End Sub
Private Sub OutBar_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
On Error GoTo ErrTrap
Select Case pLink.ObjectName
    Case "OBCustomer"
        If checkApility("FrmCustemers") = False Then
            Exit Sub
        End If
        FrmCustemers.Show
        FrmCustemers.ZOrder 0
    
    Case "OBEmployee"
        If checkApility("FrmEmployee") = False Then
            Exit Sub
        End If
        FrmEmployee.Show
        FrmEmployee.ZOrder 0
    
    Case "OBSuplier"
        If checkApility("FrmCompany") = False Then
            Exit Sub
        End If
        FrmCompany.Show
        FrmCompany.ZOrder 0
    
    Case "OBGroup"
        If checkApility("FrmGroups") = False Then
            Exit Sub
        End If
        FrmGroups.Show
        FrmGroups.ZOrder 0
    
    Case "OBItems"
        If checkApility("FrmItems") = False Then
            Exit Sub
        End If
        FrmItems.Show
        FrmItems.ZOrder 0
    
    Case "OBPriceList"
        If checkApility("FrmMainPriceList") = False Then
            Exit Sub
        End If
        FrmMainPriceList.Show
        FrmMainPriceList.ZOrder 0
    
    Case "OBSall"
        If checkApility("FrmSaleBill") = False Then
            Exit Sub
        End If
        FrmSaleBill.Show
        FrmSaleBill.ZOrder 0
    
    Case "OBPurchase"
        If checkApility("FrmBillBuy") = False Then
            Exit Sub
        End If
        FrmBillBuy.Show
        FrmBillBuy.ZOrder 0
    
    
    Case "OBReturn"
        If checkApility("FrmReturnpurchases") = False Then
            Exit Sub
        End If
        FrmReturnpurchases.Show
        FrmReturnpurchases.ZOrder 0
    
    
    Case "OBMaintenence"
        If checkApility("FrmMaintenence") = False Then
            Exit Sub
        End If
        FrmMaintenence.Show
        FrmMaintenence.ZOrder 0
    Case "OBBalance"
        If checkApility("FrmOpeningBalance") = False Then
            Exit Sub
        End If
        FrmOpeningBalance.Show
        FrmOpeningBalance.ZOrder 0
    Case "OBExpenses"
        If checkApility("FrmExpenses") = False Then
            Exit Sub
        End If
        FrmExpenses.Show
        FrmExpenses.ZOrder 0
    Case "OBCashing"
        If checkApility("FrmCashing") = False Then
            Exit Sub
        End If
        FrmCashing.Show
        FrmCashing.ZOrder 0
    Case "OBReport"
        If checkApility("FrmReports") = False Then
            Exit Sub
        End If
        FrmReports.Show
        FrmReports.ZOrder 0
    Case "OBDailyReport"
        If checkApility("FrmDailtyReport") = False Then
            Exit Sub
        End If
        FrmDailtyReport.Show
        FrmDailtyReport.ZOrder 0
    Case "ShowPrice"
        If checkApility("FrmShowPrice") = False Then
            Exit Sub
        End If
        FrmShowPrice.Show
        FrmShowPrice.ZOrder 0
    Case "Template"
        If checkApility("FrmTemplate") = False Then
            Exit Sub
        End If
        FrmTemplate.Show
        FrmTemplate.ZOrder 0
    Case "Payments"
        If checkApility("FrmPayments") = False Then
            Exit Sub
        End If
        FrmPayments.Show
        FrmPayments.ZOrder 0
    Case "OBPremium"
        FrmMkafea.Show
        FrmMkafea.ZOrder 0
    Case "OBdISCOUNT"
        FrmKhsm.Show
        FrmKhsm.ZOrder 0
    Case "OBComingTime"
        FrmPresentTime.Show
        FrmPresentTime.ZOrder 0
    Case "OBGoTime"
        FrmGoTime.Show
        FrmGoTime.ZOrder 0
    Case "AbsentRecord"
        FrmAbsent.Show
        FrmAbsent.ZOrder 0
    Case "EmpSalary"
        FrmEmpSalary.Show
        FrmEmpSalary.ZOrder 0
End Select
Exit Sub
ErrTrap:
End Sub
