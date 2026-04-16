VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmOurBarPane 
   BorderStyle     =   0  'None
   Caption         =   "‘—Ìÿ «·√Œ ’«—« "
   ClientHeight    =   6615
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   2190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   2280
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PicOutBar 
      Height          =   6645
      Left            =   30
      RightToLeft     =   -1  'True
      ScaleHeight     =   6585
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   30
      Width           =   2115
      Begin DXSIDEBARLibCtl.dxSideBar OutBar 
         Height          =   7380
         Left            =   90
         OleObjectBlob   =   "FrmOurBarPane.frx":0000
         TabIndex        =   1
         Top             =   60
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmOurBarPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim BGround As ClsBackGroundPic

    Set SysOutBar = Me.OutBar
    Set BGround = New ClsBackGroundPic
    DatName = App.path & "\Panel.txt"
    Me.LoadInterface SystemOptions.UserInterface
    OutBar.DefaultBkGround = BGround.Picture
    OutBar.ChangingSpeed = 0
    ModOutBar.LoadOutBarData SysOutBar

End Sub

Private Sub OutBar_OnDragEnterGroup(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, _
                                    ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, _
                                    ByVal GroupIndex As Integer, _
                                    ByVal ItemLinkIndex As Integer, _
                                    allow As Boolean, _
                                    OpenGroup As Boolean)
    
    If Not pLink Is Nothing Then
        '    If pLink.FileName <> "" Then
        '        If IsOutBarExistItem("User" & pLink.FileName) Then
        '            Allow = False
        '        End If
        '    End If
        Debug.Print pLink.Caption
        Debug.Print pLink.FileName
    End If

End Sub

Private Sub OutBar_OnDropFile(ByVal pFileLink As DXSIDEBARLibCtl.IdxItemLink)

    pFileLink.Caption = pFileLink.Item.Caption
    pFileLink.DefaultCaption = False
    pFileLink.Item.Caption = pFileLink.FileName
    pFileLink.Item.ObjectName = "User" & pFileLink.FileName
    pFileLink.Item.UserData = OutBar.Items.count - 1

    'Dim PName As String, FName As String
    'Dim NumImage
    'Dim Msg As String
    '
    'On Error GoTo ErrHandler
    'With Me.Cmdlg
    '
    '    .Filter = "Programs|*.exe;*.com;*.bat;*.lnk|Links|*.url|Spreadsheets|*.xls|Documents|*.doc|All files|*.*"
    '    .Flags = cdlOFNHideReadOnly
    '    .ShowOpen
    '    If .FileName <> "" Then
    '        File_Path_Name .FileName, PName, FName
    '        If IsOutBarExistItem("User" & PName & FName) = True Then
    '            Msg = "ÌÊÃœ ≈Œ ’«— „ÊÃÊœ „”»ﬁ« ·Â–« «·„·› ..!!"
    '            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    '        Set Aitem = OutBar.Items.Add
    '        NumImage = AddIcon(PName & FName)
    '        Aitem.Caption = PName & FName
    '        Aitem.UserData = OutBar.Items.count - 1
    '        Aitem.ItemLargeImage = NumImage(0)
    '        Aitem.ItemSmallImage = NumImage(1)
    '        Aitem.ObjectName = "User" & PName & FName
    '        Set Alink = Agroup.Links.Add
    '        Alink.Caption = FName
    '        Alink.DefaultCaption = False
    '        Set Alink.Item = Aitem
    '        OutBar.EditItemLinkCaption Alink
    '    End If
    'End With
End Sub

Private Sub OutBar_OnMouseDown(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)

    If Button = vbRightButton Then
        Set Alink = OutBar.GetItemLinkAt(X, Y)
        Set Agroup = OutBar.ActiveGroup

        If Not Alink Is Nothing Then
            'PopupMenu mnuLinkMenu
            Exit Sub
        End If

        Set Agroup = Nothing

        Select Case OutBar.GetFocusType(X, Y)

            Case F_AREA
                Set Agroup = OutBar.ActiveGroup

            Case F_GROUP
                Set Agroup = OutBar.GetGroupAt(X, Y)
        End Select

        If Not Agroup Is Nothing Then
            Dim i As Byte
            'mnuGroup(3).Enabled = outbar.Groups.count > 1
            'mnuGroup(5).Checked = Agroup.ItemsStyle = LargeIcon
            'For i = 0 To 2
            '    mnuBG(i).Checked = Agroup.UserData = i
            'Next
            'Visible_mnuGroup (True)
            'PopupMenu mnuGroupMenu
        End If
    End If

End Sub

Private Sub Form_Resize()
    Me.PicOutBar.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    Me.OutBar.Move Me.PicOutBar.ScaleLeft, Me.PicOutBar.ScaleTop, Me.PicOutBar.ScaleWidth, Me.PicOutBar.ScaleHeight
End Sub

Private Sub OutBar_OnMouseUp(ByVal Button As Integer, _
                             ByVal Shift As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)

    If Button = vbRightButton Then
        If Me.OutBar.Groups(0).ItemsStyle = SmallIcon Then
            mdifrmmain.MnuOutBarStyle(0).Checked = True
            mdifrmmain.MnuOutBarStyle(1).Checked = False
        Else
            mdifrmmain.MnuOutBarStyle(0).Checked = False
            mdifrmmain.MnuOutBarStyle(1).Checked = True
        End If

        If UserGroup(Agroup) Then
            mdifrmmain.MnuOutBarGroup(1).Enabled = True
            mdifrmmain.MnuOutBarGroup(2).Enabled = True
            mdifrmmain.MnuOutBarGroup(3).Enabled = True
        Else
            mdifrmmain.MnuOutBarGroup(1).Enabled = False
            mdifrmmain.MnuOutBarGroup(2).Enabled = False
            mdifrmmain.MnuOutBarGroup(3).Enabled = False
        End If

        '----------------------------------
        mdifrmmain.MnuOutBarGroup(5).Enabled = False
        mdifrmmain.MnuOutBarGroup(6).Enabled = False

        If Not Alink Is Nothing Then
            If UserItem(Alink.Item) = True Then
                mdifrmmain.MnuOutBarGroup(5).Enabled = True
                mdifrmmain.MnuOutBarGroup(6).Enabled = True
            End If
        End If

        '------------------------------------
        Me.PopupMenu mdifrmmain.MnuOutBarOptions
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ModOutBar.SaveOurBarData
End Sub

Private Sub OutBar_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, _
                                   ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, _
                                   ByVal GroupIndex As Integer, _
                                   ByVal ItemLinkIndex As Integer)
    Dim xTempItem As DXSIDEBARLibCtl.dxItem
    Dim Msg As String

    On Error GoTo ErrTrap
    Set xTempItem = pLink.Item

    If UserItem(xTempItem) = True Then
        OpenFile pLink.Item.Caption
    Else

        Select Case pLink.ObjectName

                '--------------------------------------------------
                '»œ«Ì… „Ã„Ê⁄… «·»Ì«‰«  «·√”«”Ì…
                '--------------------------------------------------
            Case "OBCustomer"
                '»Ì«‰«  «·⁄„·«¡
                OpenScreen CustomersScreen

            Case "OBEmployee"
                '»Ì«‰«  «·„ÊŸ›Ì‰
                OpenScreen EmployeesScreen

            Case "OBSuplier"
                '»Ì«‰«  «·„Ê—œÌ‰
                OpenScreen SuppliersScreen

            Case "dxItemLink2"
                '»Ì«‰«  ‘—ﬂ«  «·’Ì«‰…
                OpenScreen ManCompaniesScreen

            Case "OBGroup"
                '»Ì«‰«  „Ã„Ê⁄«  «·√’‰«›
                OpenScreen ItemsGroupsScreen

            Case "OBItems"
                '»Ì«‰«  «·√’‰«›
                OpenScreen ItemsDataScreen

                '--------------------------------------------------
                '»œ«Ì… „Ã„Ê⁄… «·„⁄«„·«  «· Ã«—Ì…
                '--------------------------------------------------
            Case "ShowPrice"
                '‘«‘… ⁄—Ê÷ «·√”⁄«—
                OpenScreen ScreensName.ShowPriceScreen

            Case "Template"
                '‘«‘… «·⁄—Ê÷ «·Ã«Â“…
                OpenScreen TemplateScreen

            Case "OBSall"
                '‘«‘… ›« Ê—… «·»Ì⁄
                OpenScreen InvoiceScreen

            Case "OBPurchase"
                '‘«‘… ›« Ê—… «·„‘ —Ì« 
                OpenScreen PurchaseScreen

            Case "dxItemLink21"
                '‘«‘… „— Ã⁄ «·„»Ì⁄« 
                OpenScreen RetrunSalles

            Case "OBReturn"
                '‘«‘… „— Ã⁄ «·„‘ —Ì« 
                OpenScreen RetrunPurchse

            Case "OBBalance"
                '«·—’Ìœ «·≈›  «ÕÏ ··„Œ«“‰
                OpenScreen OpenStockBalance

            Case "OBPriceList"
                '‘«‘… ﬁ«∆„… «·√”⁄«—
                OpenScreen ItemsMainPriceLise

            Case "OBMaintenence"

                If checkApility("FrmManStore") = False Then
                    Exit Sub
                End If

                'FrmManStore.show
                'FrmManStore.ZOrder 0

                '-----------------------------------------------
                '„Ã„Ê⁄… «·„⁄«„·«  «·„«·Ì…
                '-----------------------------------------------
            Case "OBExpenses"
                '‘«‘… »Ì«‰«  «·„’—Ê›« 
                OpenScreen ExpensesDataScreen

            Case "Payments"
                '‘«‘… »Ì«‰«  «·„œ›Ê⁄« 
                OpenScreen PaymentsDataScreen

            Case "OBCashing"
                '‘«‘… »Ì«‰«  «·„ﬁ»Ê÷« 
                OpenScreen CashingDataScreen

            Case "dxItemLink1"
                '‘«‘… «·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…
                OpenScreen AllowsDiscountsScreen

            Case "dxItemLink10"
                ' ‰»ÌÂ «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…
                OpenScreen PopUpShowPaymentTime

            Case "dxItemLink3"
                ' Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ
                OpenScreen ReceiptPartScreen

            Case "dxItemLink11"
                '‘«‘… «·√ﬁ”«ÿ «·„” Õﬁ…
                OpenScreen PopUpShowInstallmentMustPay

            Case "OBReport"

                If checkApility("FrmReports") = False Then
                    Exit Sub
                End If

                FrmReports.show
                FrmReports.ZOrder 0

            Case "OBDailyReport"

                If checkApility("FrmDailtyReport") = False Then
                    Exit Sub
                End If

                FrmDailtyReport.show
                FrmDailtyReport.ZOrder 0

            Case "OBPremium"
                FrmMkafea.show
                FrmMkafea.ZOrder 0

            Case "OBdISCOUNT"
                FrmKhsm.show
                FrmKhsm.ZOrder 0

            Case "OBComingTime"
              ' FrmPresentTime.show
              '  FrmPresentTime.ZOrder 0

            Case "OBGoTime"
         '       FrmGoTime.show
         '       FrmGoTime.ZOrder 0

            Case "AbsentRecord"
            '    FrmAbsent.show
            '    FrmAbsent.ZOrder 0

            Case "EmpSalary"
             '   FrmEmpSalary.show
             '   FrmEmpSalary.ZOrder 0

            Case "dxItemLink5"
            '    FrmChiqueRelease.show

            Case "dxItemLink6"

                If checkApility("FrmBoxDeposit") = False Then
                    Exit Sub
                End If

                'FrmBoxDeposit.show
                'FrmBoxDeposit.ZOrder 0

            Case "dxItemLink7"

                If checkApility("FrmBoxDrawing") = False Then
                    Exit Sub
                End If

                FrmBoxDrawing.show
                FrmBoxDrawing.ZOrder 0

            Case "dxItemLink8"
                FrmBoxIncapacity.show

                '----------------------------------------------
                'Ã“¡ «·≈” ⁄·«„« 
                '----------------------------------------------
            Case "dxItemLink15"
                '«·≈” ⁄·«„ ⁄‰ √—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰
                OpenScreen PopUpShowCusBalances

            Case "dxItemLink12"
                '‘«‘… «·√” ⁄·«„ ⁄‰ «—’œ… «·Œ“‰… «·¬‰
                OpenScreen PopUpShowBoxesAccounts

            Case "dxItemLink13"
                OpenScreen PopUpShowItemsRequest

            Case "dxItemLink9"
                OpenScreen CheckItemSerial

            Case "dxItemLink14"
                OpenScreen PopUpShowItemQuantity

            Case "dxItemLink4"
                OpenScreen PopUpShowGuaranteeAlram

            Case "dxItemLink17"
                OpenScreen PopUpSowStagnantItems

            Case "dxItemLink18"
                OpenScreen StatisticsShow

            Case "dxItemLink19"

                '            If SystemOptions.SysDataBaseType = AccessDataBase Then
                '                Msg = "Â–Â «·√„ﬂ«‰Ì… „ «Õ… ›ﬁÿ ›Ï ‰”Œ… «·‘»ﬂ«  „‰ »—‰«„Ã «·„Õ«”» «·⁄—»Ï"
                '                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '                Exit Sub
                '            End If
                '            FrmOlapShow.Show
            Case "dxItemLink20"
                OpenScreen PopUpShowStockMovement

            Case "dxItemLink22"
     '           FrmManGoBack.show

            Case "dxItemLink16"
      '          FrmMaintenenceGoIn.show

            Case "dxItemLink23"
'                FrmManCusRecive.show

            Case "dxItemLink24"
'                FrmManStoreStock.show

            Case "dxItemLink25"
           '     FrmManAlram.show vbModal

            Case "dxItemLink26"
                '‘«‘… „ Ê”ÿ  ﬂ·›… «·’‰›
                OpenScreen PopUpShowItemCostScreen

            Case Else
                Debug.Print "X" & pLink.ObjectName
        End Select

    End If

    'Exit Sub
ErrTrap:
End Sub

Public Sub LoadInterface(IntInterface As SystemInterface)
    Dim XPanel As MSComctlLib.Panel
    Dim X As DXSIDEBARLibCtl.IconStyle
    Dim i As Integer

    If SystemOptions.SysMantainceAllow = True Then
        Me.OutBar.Groups(3).Visible = True
    Else
        Me.OutBar.Groups(3).Visible = False
    End If

    Screen.MousePointer = vbArrowHourglass
    X = GetSetting(StrAppRegPath, "OutBarOptions", "ItemsStyle", X)

    For i = 0 To Me.OutBar.Groups.count - 1
        Me.OutBar.Groups(i).ItemsStyle = X
    Next i

    If IntInterface = ArabicInterface Then
        Me.RightToLeft = True

        With Me.OutBar
            .GroupAlignment = egaRight
            .GroupsFont.Bold = True
            .ItemsFont.Bold = False
            .Groups(0).Caption = "»Ì«‰«  √”«”Ì…"
            .Groups(0).Links(0).Item.Caption = "«·„ÊŸ›Ì‰"
            .Groups(0).Links(1).Item.Caption = "«·⁄„·«¡"
            .Groups(0).Links(2).Item.Caption = "«·„Ê—œÌ‰"
            .Groups(0).Links(3).Item.Caption = "‘—ﬂ«  «·’Ì«‰…"
            .Groups(0).Links(4).Item.Caption = "«·„Ã„Ê⁄« "
            .Groups(0).Links(5).Item.Caption = "«·√’‰«›"
            
            .Groups(1).Caption = "«·„⁄«„·«  «· Ã«—Ì…"
            .Groups(1).Links(0).Item.Caption = "ﬁ«∆„… «·√”⁄«—"
            .Groups(1).Links(1).Item.Caption = "⁄—÷ √”⁄«—"
            .Groups(1).Links(2).Item.Caption = "⁄—Ê÷ Ã«Â“…"
            .Groups(1).Links(3).Item.Caption = "»Ì⁄"
            .Groups(1).Links(4).Item.Caption = "‘—«¡"
            .Groups(1).Links(5).Item.Caption = "„— Ã⁄ «·„»Ì⁄« "
            .Groups(1).Links(6).Item.Caption = "„— Ã⁄ „‘ —Ì« "
            .Groups(1).Links(7).Item.Caption = "—’Ìœ ≈›  «ÕÏ"
             
            .Groups(2).Caption = "«·„⁄«„·«  «·„«·Ì…"
            .Groups(2).Links(0).Item.Caption = "«·„’—Ê›« "
            .Groups(2).Links(1).Item.Caption = "«·„œ›Ê⁄« "
            .Groups(2).Links(2).Item.Caption = "«·„ﬁ»Ê÷« "
            .Groups(2).Links(3).Item.Caption = "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…"
            .Groups(2).Links(4).Item.Caption = " Õ’Ì· Ê”œ«œ √ﬁ”«ÿ"
            .Groups(2).Links(5).Item.Caption = " Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
            .Groups(2).Links(6).Item.Caption = "≈Ìœ«⁄ ›Ï «·Œ“‰…"
            .Groups(2).Links(7).Item.Caption = "”Õ» „‰ «·Œ“‰…"
            .Groups(2).Links(8).Item.Caption = "“Ì«œ… Ê⁄Ã“ ›Ï ‰ﬁœÌ… «·Œ“‰…"
            .Groups(3).Caption = "«·’Ì«‰… Ê«·÷„«‰« "
            .Groups(3).Links(0).Item.Caption = " ‰»ÌÂ ﬁ”„ «·’Ì«‰…"
            .Groups(4).Caption = "‘∆Ê‰ «·„ÊŸ›Ì‰"
            .Groups(4).Links(0).Item.Caption = "«·„ﬂ«›¬ "
            .Groups(4).Links(1).Item.Caption = "«·Œ’Ê„« "
            .Groups(4).Links(2).Item.Caption = "„Ê«⁄Ìœ «·Õ÷Ê—"
            .Groups(4).Links(3).Item.Caption = "„Ê«⁄Ìœ «·√‰’—«›"
            .Groups(4).Links(4).Item.Caption = " ”ÃÌ· «·€Ì«»"
            .Groups(4).Links(5).Item.Caption = "—Ê« » «·„ÊŸ›Ì‰"
        
            .Groups(5).Caption = "«· ﬁ«—Ì—"
            .Groups(5).Links(0).Item.Caption = "„œÌ— «· ﬁ«—Ì—"
            .Groups(5).Links(1).Item.Caption = " ﬁ—Ì— «·ÌÊ„"
            .Groups(6).Caption = "≈” ⁄·«„« "
            .Groups(6).Links(0).Item.Caption = "√—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰"
            .Groups(6).Links(1).Item.Caption = "√Ê—«ﬁ „«·Ì… „” Õﬁ…"
            .Groups(6).Links(2).Item.Caption = "«·√ﬁ”«ÿ «·„ÿ·Ê»…"
            .Groups(6).Links(3).Item.Caption = "—’Ìœ «·Œ“‰… «·√‰"
            .Groups(6).Links(4).Item.Caption = "«·√’‰«› «·„ÿ·Ê»…"
            .Groups(6).Links(5).Item.Caption = "»ÕÀ ⁄‰ ”Ì—Ì«·"
            .Groups(6).Links(6).Item.Caption = "ﬂ„Ì… ’‰›"
            .Groups(6).Links(7).Item.Caption = "÷„«‰ «·√’‰«›"
            .Groups(6).Links(8).Item.Caption = "«·√’‰«› «·—«ﬂœ…"
            .Groups(6).Links(9).Item.Caption = "Õ—ﬂ… «·„Œ“Ê‰"
            .Groups(6).Links(10).Item.Caption = "Õ—ﬂ…  ﬂ·›… ’‰›"
        End With
    
    ElseIf IntInterface = EnglishInterface Then
        Me.RightToLeft = False

        With Me.OutBar
            .GroupAlignment = egaLeft
            .GroupsFont.Bold = True
            .ItemsFont.Bold = True
         
            .Groups(0).Caption = "Basic Data"
            .Groups(0).Links(0).Item.Caption = "Employess"
            .Groups(0).Links(1).Item.Caption = "Customers"
            .Groups(0).Links(2).Item.Caption = "Suppliers"
            .Groups(0).Links(3).Item.Caption = "Maintenance Companies"
            .Groups(0).Links(4).Item.Caption = "Items Groups"
            .Groups(0).Links(5).Item.Caption = "Items"
             
            .Groups(1).Caption = "Inventory"
            .Groups(1).Links(0).Item.Caption = "Price List"
            .Groups(1).Links(1).Item.Caption = "Price Order"
            .Groups(1).Links(2).Item.Caption = "Template Price"
            .Groups(1).Links(3).Item.Caption = "Bill Invoice"
            .Groups(1).Links(4).Item.Caption = "Purchase Invoice"
            .Groups(1).Links(5).Item.Caption = "Return Sales"
            .Groups(1).Links(6).Item.Caption = "Purchase Retruns"
            .Groups(1).Links(7).Item.Caption = "Beginning Stock Balance"""
             
            .Groups(2).Caption = "Financial"
            .Groups(2).Links(0).Item.Caption = "Expenses"
            .Groups(2).Links(1).Item.Caption = "Notes Payable"
            .Groups(2).Links(2).Item.Caption = "Notes Receivable"
            .Groups(2).Links(3).Item.Caption = "Allowed and acquired Discounts"
            .Groups(2).Links(4).Item.Caption = "Getting Installment"
            .Groups(2).Links(5).Item.Caption = "Check Release"
            .Groups(2).Links(6).Item.Caption = "Box Deposit"
            .Groups(2).Links(7).Item.Caption = "Box Drawing"
            .Groups(2).Links(8).Item.Caption = "Box Incapacity && Increase"
        
            .Groups(3).Caption = "Maintenance"
            .Groups(3).Links(0).Item.Caption = "Maintenance Alarm"
            .Groups(3).Links(1).Item.Caption = "Maintenance Store"
            .Groups(3).Links(2).Item.Caption = "Maintenance Query"
              
            .Groups(4).Caption = "Empolyees Mangment"
            .Groups(4).Links(0).Item.Caption = "Premium"
            .Groups(4).Links(1).Item.Caption = "Punishment"
            .Groups(4).Links(2).Item.Caption = "Presence Recording"
            .Groups(4).Links(3).Item.Caption = "Departure Recording"
            .Groups(4).Links(4).Item.Caption = "Absence Recording"
            .Groups(4).Links(5).Item.Caption = "Monthly Payroll"
        
            .Groups(5).Caption = "Reports"
            .Groups(6).Links(0).Item.Caption = "General Reports"
            .Groups(5).Links(1).Item.Caption = "Daily Reports"
            .Groups(6).Caption = "Information"
            .Groups(6).Links(0).Item.Caption = "Customers Balance"
            .Groups(6).Links(1).Item.Caption = "Due Notes"
            .Groups(6).Links(2).Item.Caption = "Due Installments"
            .Groups(6).Links(3).Item.Caption = "Current Box Accout"
            .Groups(6).Links(4).Item.Caption = "Required Items"
            .Groups(6).Links(5).Item.Caption = "Search For Item Serial"
            .Groups(6).Links(6).Item.Caption = "Item Stock"
            .Groups(6).Links(7).Item.Caption = "Items Guarantee"
            .Groups(6).Links(8).Item.Caption = "Stagnant Items"
            .Groups(6).Links(9).Item.Caption = "Stock Cost & Movement"
            .Groups(6).Links(10).Item.Caption = "Item Stock Cost"
        End With

    End If

    Screen.MousePointer = vbDefault
End Sub

