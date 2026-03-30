VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPrintBarcode 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17505
   Icon            =   "FrmPrintBarcode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   17505
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   1305
   End
   Begin VB.CheckBox chkhalf 
      Alignment       =   1  'Right Justify
      Caption         =   "‰’ «·ﬂ„ÌÂ"
      Height          =   255
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox Check17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ÕœÌœ / «·€«¡ «·ﬂ·"
      Height          =   195
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5070
      Picture         =   "FrmPrintBarcode.frx":038A
      RightToLeft     =   -1  'True
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   5550
      Width           =   255
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   4515
      Left            =   45
      TabIndex        =   0
      Top             =   960
      Width           =   17370
      _cx             =   30639
      _cy             =   7964
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   28
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmPrintBarcode.frx":0714
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   5880
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ButtonImage     =   "FrmPrintBarcode.frx":0B07
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmPrintBarcode.frx":0EA1
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   5880
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmPrintBarcode.frx":123B
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„Ê—œ"
      Height          =   255
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.Label LblCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì„ﬂ‰ﬂ  ÕœÌœ «·√’‰«› «· Ì  —€» ›Ì ÿ»«⁄ Â« „‰ «·⁄„Êœ ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5550
      Width           =   4905
   End
   Begin VB.Label LblID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1830
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "  ÿ»«⁄… »«—ﬂÊœ ··√’‰«›  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Index           =   4
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   17565
   End
End
Attribute VB_Name = "FrmPrintBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = False
            Next i

        End With

    End If

 

End Sub
Private Sub ChangeLang()
    LblCaption(4).Caption = "Barcode printing for varieties"
    Me.Caption = "Barcode printing"
    LblCaption(0).Caption = "You can select the items you want to print from the Print column"
    ISButton1.Caption = "Print"
    CmdExit.Caption = "Exit"
    Check17.Caption = "Select / Cancel All"
   With Me.FG

        .TextMatrix(0, .ColIndex("Print")) = " Print "
        .TextMatrix(0, .ColIndex("Code")) = "Code "
        .TextMatrix(0, .ColIndex("barcodeno")) = "Barcode "
        .TextMatrix(0, .ColIndex("Name")) = "Item Name"
        .TextMatrix(0, .ColIndex("PartNo")) = "Part No "
        .TextMatrix(0, .ColIndex("Cost")) = " Sales Price"
        .TextMatrix(0, .ColIndex("VatYou")) = "Vat ratio "
        .TextMatrix(0, .ColIndex("VAT")) = "Vat  "
        .TextMatrix(0, .ColIndex("Total")) = "Total  "
        .TextMatrix(0, .ColIndex("Qty")) = "Qty  "
        .TextMatrix(0, .ColIndex("LotNO")) = "LotNO  "
        .TextMatrix(0, .ColIndex("ProductionDate")) = "Production Date  "
        .TextMatrix(0, .ColIndex("ExpiryDate")) = "Expiry Date"

    End With


End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim RowNum As Integer
    Dim ItemCount As Integer
    'cBarcode.AddItem
    cBarcode.ClearItems

    For RowNum = 1 To FG.rows - 1

        If FG.cell(flexcpChecked, RowNum, FG.ColIndex("Print")) = flexChecked Then
            If Not IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))) Then

                For ItemCount = 1 To val(FG.TextMatrix(RowNum, FG.ColIndex("Qty")))
                    cBarcode.AddItem FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), FG.TextMatrix(RowNum, FG.ColIndex("Name")) & "/" & FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Cost"))
                Next ItemCount

            End If
        End If

    Next RowNum

    'cBarcode.AddItem FG.TextMatrix(2, FG.ColIndex("Code")), "yasser", "ahmed"
'    FrmSetting.show vbModal

End Sub

Public Sub DBCboClientName_Click(Area As Integer)
     On Error Resume Next
    Dim Fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 2
    TxtSearchCode.text = Fullcode

End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col = FG.ColIndex("Name") Or Col = FG.ColIndex("Code") Then
        Cancel = True
    End If

End Sub

Private Sub Form_Activate()
    'LodaData
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As New ClsBackGroundPic
      Dim Dcombos As New ClsDataCombos
 
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True


    CenterForm Me

    FormPostion Me, GetPostion
    If SystemOptions.UserInterface = EnglishInterface Then
        ChangeLang
    End If
    
    LoadIcon

    Check17.value = vbChecked
  Set FG.WallPaper = BGround.Picture
  
    '    Exit Sub
        
ErrTrap:
End Sub

Private Sub LoadIcon()
    On Error GoTo ErrTrap

    With FG
        .cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .cell(flexcpPicture, 0, .ColIndex("Code")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .cell(flexcpPicture, 0, .ColIndex("Cost")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .cell(flexcpPicture, 0, .ColIndex("Qty")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .cell(flexcpPicture, 0, .ColIndex("Print")) = mdifrmmain.ImgLstTree.ListImages("Print").Picture
        .cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub LodaData()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim RowNum As Integer
    Dim Percentg As Double
Dim cCompanyInfo As New ClsCompanyInfo
    If LblID.Caption = "" Then Exit Sub
  '  StrSQL = "SELECT * FROM QryBarcode WHERE Transaction_ID=" & val(LblID.Caption)
  'StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.Quantity, "
  'StrSQL = StrSQL & "   dbo.TblItems.SallingPrice , dbo.TblItems.PartNo, dbo.TblItems.ItemNamee, dbo.TblItems.barCodeNO"
'StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'StrSQL = StrSQL & "  dbo.TblItems INNER JOIN"
'StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'StrSQL = StrSQL & "  WHERE Transactions.Transaction_ID=" & val(lblid.Caption) StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.Quantity, "
    
    
    
    
  StrSQL = " SELECT     TOP 100 PERCENT dbo.Transaction_Details.ProductionDate, dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO, "
    StrSQL = StrSQL & "                  dbo.Transactions.TransactionComment, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Item_ID,CAST (  Transactions.InvoiceOrderNo AS NVARCHAR(10)) AS InvoiceOrderNo, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.ItemCase,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Price,Emp4.Emp_Name as TechName, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.Transaction_Details.ColorID,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblUnites.UnitName, dbo.TblItems.SallingPrice, dbo.TblItems.PartNo, dbo.Transaction_Details.ShowQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.showPrice, dbo.Transactions.NoteSerial1, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                  dbo.TblCustemers.CusNamee, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStore.StoreNamee,"
    StrSQL = StrSQL & "                  dbo.TblItems.ItemNamee, dbo.TblUnites.UnitNamee, dbo.TblCustemers.Fullcode AS CusSupCode, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode AS EmpFullCode, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    StrSQL = StrSQL & "                  dbo.Transactions.FixesAssetsID, dbo.FixedAssets.Name AS FaName, dbo.FixedAssets.Fullcode AS FaCode, dbo.FixedAssets.namee AS FaNameE,"
    StrSQL = StrSQL & "                  dbo.TblStore.Code AS StoreCode, dbo.Transactions.CashCustomerName, dbo.Transactions.Nots2, dbo.Transactions.order_no, dbo.Transactions.BillBasedOn,"
    StrSQL = StrSQL & "                  dbo.Transactions.ManualNo1, dbo.Transactions.ManualNo2, dbo.Transactions.ManualNO, dbo.Transaction_Details.MixNo, dbo.Transaction_Details.MaxQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.TypeVAT, dbo.Transaction_Details.ISSUEDQTY, dbo.Transaction_Details.TotalPriceNoHours, dbo.Transaction_Details.TotalInvoiceQty,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.PriceNoHours, dbo.Transaction_Details.NoHours, dbo.Transaction_Details.Vat, dbo.Transaction_Details.Vatyo,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.IsExpirDate, dbo.Transaction_Details.L, dbo.Transaction_Details.W, dbo.Transaction_Details.H1, dbo.Transaction_Details.H2,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.OrderNo, dbo.Transaction_Details.StoreID2, TblStore_1.StoreName AS Expr1, TblStore_1.StoreNamee AS Expr2,"
    '''''''''''''''''''''''''''Ship
    StrSQL = StrSQL & "                  dbo.Transaction_Details.QtyFaqtors "
    StrSQL = StrSQL & "  ,  Transactions.ShipOrderNo, "
    StrSQL = StrSQL & "       Transactions.ShipEnquieryNo, "
    StrSQL = StrSQL & "       Transactions.ShipAccountNo, "
    StrSQL = StrSQL & "       Transactions.ShipCustomerName, "
    StrSQL = StrSQL & "       Transactions.ShipDistance, "
    StrSQL = StrSQL & "       Transactions.ShipArea, "
    StrSQL = StrSQL & "       Transactions.ShipSiteNo, "
    StrSQL = StrSQL & "       Transactions.ShipProjectName, "
    StrSQL = StrSQL & "       Transactions.ShipStructuralElement, "
    StrSQL = StrSQL & "       Transactions.ShipMixDescription, "
    StrSQL = StrSQL & "       Transactions.ShipDriverName, "
    StrSQL = StrSQL & "       Transactions.ShipPipeLine, "
    StrSQL = StrSQL & "       Transactions.ShipPump, "
    StrSQL = StrSQL & "       Transactions.ShipTruckNo, "
    StrSQL = StrSQL & "       Transactions.ShipIceTemp, "
    StrSQL = StrSQL & "       Transactions.ShipTotalDeleveryd, "
    StrSQL = StrSQL & "       Transactions.ShipThisLoad, "
    StrSQL = StrSQL & "       Transactions.ShipDayOrder, "
    StrSQL = StrSQL & "       Transactions.ShipTripNo, "
    StrSQL = StrSQL & "       Transactions.ShipPlantNo, "
    StrSQL = StrSQL & "       Transactions.ShipBatched, "
    StrSQL = StrSQL & "       Transactions.ShipRestunedPlant, "
    StrSQL = StrSQL & "       Transactions.ShipEndDischarge, "
    StrSQL = StrSQL & "       Transactions.ShipStartDisCharge, "
    StrSQL = StrSQL & "       Transactions.ShipOnSite,Transaction_Details.Remarks"
    '*************************************
    StrSQL = StrSQL & "   FROM         dbo.TblStore INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore TblStore_1 ON dbo.Transaction_Details.StoreID2 = TblStore_1.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmpDepartments ON dbo.Transactions.DepartementID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    
    StrSQL = StrSQL & "                  Left Outer join TblEmployee Emp4 On dbo.Transaction_Details.EmpId4 = Emp4.Emp_ID "
    
    
    
    
    
StrSQL = " SELECT Transaction_Details.NoCount   ,dbo.Transaction_Details.ProductionDate, TblUnites.UnitName,Transaction_Details.LineId,TblUnites.UnitID,TblUnites.UnitNamee, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.Quantity, "
StrSQL = StrSQL & " dbo.TblItems.PartNo, dbo.TblItems.ItemNamee, dbo.TblItems.barCodeNO, dbo.TblItemsUnits.UnitSalesPrice AS SallingPrice, dbo.Transaction_Details.UnitId,"

StrSQL = StrSQL & "                  dbo.Transaction_Details.NoCount, dbo.Transaction_Details.Area, Transaction_Details.Length,Transaction_Details.Height,Transaction_Details.Width,"
StrSQL = StrSQL & "                  dbo.Transaction_Details.IsExpirDate, dbo.Transaction_Details.L, dbo.Transaction_Details.W, dbo.Transaction_Details.H1, dbo.Transaction_Details.H2,Transaction_Details.Remarks,"
 
StrSQL = StrSQL & "                   dbo.Transaction_Details.ColorID , "
StrSQL = StrSQL & "                  dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "                   dbo.Transactions.CusID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "  dbo.Transaction_Details.ExpiryDate, dbo.Transaction_Details.LotNO"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & " dbo.TblItems INNER JOIN"
StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & " dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & " dbo.TblItemsUnits ON dbo.Transaction_Details.UnitId = dbo.TblItemsUnits.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID "

StrSQL = StrSQL & "           Left outer join       dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN "
StrSQL = StrSQL & "                  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN "
StrSQL = StrSQL & "                  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID "

StrSQL = StrSQL & "           Left outer join       TblUnites ON dbo.Transaction_Details.UnitId  = TblUnites.UnitID"

StrSQL = StrSQL & "                  LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID "
    
StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_ID = " & val(LblID.Caption) & ")"


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (rs.EOF Or rs.BOF) Then

        With FG
            .rows = rs.RecordCount + 1

            For RowNum = 1 To rs.RecordCount
                 .TextMatrix(RowNum, .ColIndex("Item_ID")) = IIf(IsNull(rs("Item_ID").value), "", rs("Item_ID").value)
              If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
              Else
                 .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
              End If
                .TextMatrix(RowNum, .ColIndex("ProductionDate")) = IIf(IsNull(rs("ProductionDate").value), "", rs("ProductionDate").value)
                .TextMatrix(RowNum, .ColIndex("ExpiryDate")) = IIf(IsNull(rs("ExpiryDate").value), "", rs("ExpiryDate").value)
                .TextMatrix(RowNum, .ColIndex("LotNO")) = IIf(IsNull(rs("LotNO").value), "", rs("LotNO").value)
                .TextMatrix(RowNum, .ColIndex("barcodeno")) = IIf(IsNull(rs("barcodeno").value), "", rs("barcodeno").value)
                .TextMatrix(RowNum, .ColIndex("Code")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(RowNum, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
                
                .TextMatrix(RowNum, .ColIndex("ColorID")) = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
                .TextMatrix(RowNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RowNum, .ColIndex("ColorName")) = IIf(IsNull(rs("ColorName").value), "", rs("ColorName").value)
                .TextMatrix(RowNum, .ColIndex("SizeName")) = IIf(IsNull(rs("SizeName").value), "", rs("SizeName").value)
                .TextMatrix(RowNum, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                
.TextMatrix(RowNum, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value)
.TextMatrix(RowNum, .ColIndex("UnitNamee")) = IIf(IsNull(rs("UnitNamee").value), "", rs("UnitNamee").value)
.TextMatrix(RowNum, .ColIndex("UnitID")) = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
.TextMatrix(RowNum, .ColIndex("LineId")) = IIf(IsNull(rs("LineId").value), "", rs("LineId").value)
                
                .TextMatrix(RowNum, .ColIndex("Width")) = IIf(IsNull(rs("Width").value), "", rs("Width").value)
                .TextMatrix(RowNum, .ColIndex("Height")) = IIf(IsNull(rs("Height").value), "", rs("Height").value)
                .TextMatrix(RowNum, .ColIndex("Length")) = IIf(IsNull(rs("Length").value), "", rs("Length").value)
                '.TextMatrix(RowNum, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
            
                .TextMatrix(RowNum, .ColIndex("Cost")) = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                .TextMatrix(RowNum, .ColIndex("Qty")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
                .TextMatrix(RowNum, .ColIndex("PrintedQty")) = 1 '  IIf(IsNull(rs("NoCount").value), "", rs("NoCount").value)
               If .TextMatrix(RowNum, .ColIndex("ExpiryDate")) = "" Then
               .TextMatrix(RowNum, .ColIndex("PrintedQty")) = .TextMatrix(RowNum, .ColIndex("Qty"))
               Else
              .TextMatrix(RowNum, .ColIndex("PrintedQty")) = IIf(IsNull(rs("NoCount").value), "", rs("NoCount").value)
              
              End If
              If val(.TextMatrix(RowNum, .ColIndex("PrintedQty"))) = 0 Then
               .TextMatrix(RowNum, .ColIndex("PrintedQty")) = 1
               End If
               
                
                .cell(flexcpChecked, RowNum, .ColIndex("Print")) = flexChecked
                If SystemOptions.AllItemInVAT = True Then
                  Percentg = val(cCompanyInfo.VATItems)
               Else
                  Percentg = PercentgValueAddedBarcode(Date, val(.TextMatrix(RowNum, .ColIndex("Item_ID"))), 21)
               End If
             If Percentg = -1 Then
              Percentg = 0
             End If
             .TextMatrix(RowNum, .ColIndex("VatYou")) = Percentg
             If Percentg <> 0 Then
             .TextMatrix(RowNum, .ColIndex("VAT")) = Percentg * val(.TextMatrix(RowNum, .ColIndex("Cost"))) / 100
             Else
             .TextMatrix(RowNum, .ColIndex("VAT")) = 0
             End If
              .TextMatrix(RowNum, .ColIndex("Total")) = val(.TextMatrix(RowNum, .ColIndex("VAT"))) + val(.TextMatrix(RowNum, .ColIndex("Cost")))

                rs.MoveNext
            Next RowNum

        End With

    End If

    Exit Sub
ErrTrap:
End Sub
Public Function code128$(chaine$)
  'Cette fonction est rÈgie par la Licence GÈnÈrale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'ParamËtres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichÈe avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramËtre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  code128$ = ""
  If Len(chaine$) > 0 Then
  'VÈrifier si caractËres valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(mId$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intÈressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au dÈbut ou ‡ la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'DÈbuter sur table C / Starting with table C
              code128$ = CHR$(210)
            Else 'Commuter sur table C / Switch to table C
              code128$ = code128$ & CHR$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then code128$ = CHR$(209) 'DÈbuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = val(mId$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            code128$ = code128$ & CHR$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            code128$ = code128$ & CHR$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caractËre en table B / Process 1 digit with table B
          code128$ = code128$ & mId$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clÈ de contrÙle / Calculation of the checksum
      For i% = 1 To Len(code128$)
        dummy% = Asc(mId$(code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clÈ / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Ajout de la clÈ et du STOP / Add the checksum and the STOP
      code128$ = code128$ & CHR$(checksum&) & CHR$(211)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caractËres ‡ partir de i% sont numÈriques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(mId$(chaine$, i% + mini%, 1)) < 48 Or Asc(mId$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function

Function addtotable(NoOfRow As Variant, code As String, cost As Variant, Optional PartNo As String = "", Optional Name As String = "" _
, Optional NameE As String, Optional Color As String, Optional size As String, Optional Class As String, Optional lotNo As String, Optional ExpiryDate As String, Optional Item_ID As Double, Optional VatYou As Double, Optional Vat As Double, Optional total As Double, Optional ProductionDate As String, Optional Suppliercode As String, Optional SupplierName As String, Optional Qty As Double _
, Optional ColorID As String = "", Optional CusName As String = "", Optional colorname As String = "", Optional sizename As String = "" _
, Optional Remarks As String = "", Optional Width As String = "", Optional Height As String = "", Optional Length As String = "", _
Optional UnitID As Long = 0, Optional UnitName As String = "", Optional UnitNamee As String = "", Optional LineID As Long = 0)

 
 
 
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim i As Integer
    Dim Zcode As String
    Dim Zcode128 As String
    str = "select * from   TblPrintBarCode where 1=-1"
   rs.Open str, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 1 To NoOfRow
        rs.AddNew

        rs("Item_ID").value = Item_ID
        rs("PartNo").value = PartNo
        rs("Suppliercode").value = Suppliercode
        rs("Suppliername").value = SupplierName
        rs("qty").value = Qty
         
 
        'ZINA bARCODE############################################
        Zcode = mId(code, 1, 4) & mId(ExpiryDate, 1, 2) & mId(ExpiryDate, 4, 2) & mId(ExpiryDate, 9, 2) & mId(Suppliercode, 1, 3) & Qty
        Zcode = zeropadding(Zcode, 18, True)
        
         
        
        rs("Zcode").value = Zcode
        Zcode128 = code128$(Zcode)
        rs("Zcode128").value = code128$(Zcode)
        'ZINA bARCODE############################################
        
        'Suppliercode as  String,Optional  Suppliername
        
        rs("code").value = code
        rs("code128").value = code128$(code)
        
        
        rs("cost").value = val(cost)
        rs("Name").value = Name
      '  rs("NameE").value = NameE
        rs("Color").value = Color
        rs("size").value = size
        rs("class").value = Class
        
       

'rs("ColorID").value = ColorID
rs("CusName").value = CusName
rs("ColorName").value = colorname
rs("SizeName").value = sizename
rs("Remarks").value = Remarks
rs("Width").value = Width
rs("Height").value = Height

rs("UnitID").value = UnitID
rs("UnitName").value = UnitName
rs("UnitNamee").value = UnitNamee
rs("LineID").value = LineID

rs("Length").value = Length
        
        rs("LotNO").value = lotNo
        rs("VatYou").value = VatYou
        rs("VAT").value = Vat
        rs("Total").value = total
        rs("ExpiryDate").value = IIf(ExpiryDate = "", Null, ExpiryDate)
        rs("ProductionDate").value = IIf(ProductionDate = "", Null, ProductionDate)
        
        
        rs.update
    Next i
'
End Function

Private Sub ISButton1_Click()
    Dim str As String

    Dim RowNum As Integer
    Dim ItemCount As Integer
    str = "Delete  TblPrintBarCode"
    Cn.Execute str

    'cBarcode.AddItem
    ' cBarcode.ClearItems
    For RowNum = 1 To FG.rows - 1

        If FG.cell(flexcpChecked, RowNum, FG.ColIndex("Print")) = flexChecked Then
            If Not IsNull(FG.TextMatrix(RowNum, FG.ColIndex("PrintedQty"))) Then
           
           If chkhalf.value = vbChecked Then
 addtotable val(FG.TextMatrix(RowNum, FG.ColIndex("PrintedQty"))) / 2, FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), val(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))), FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Name")), , , , , FG.TextMatrix(RowNum, FG.ColIndex("LotNO")), FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")), val(FG.TextMatrix(RowNum, FG.ColIndex("Item_ID"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VatYou"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VAT"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Total"))), (FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), TxtSearchCode.text, DBCboClientName.text, val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))), _
        Trim(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("CusName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("ColorName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("SizeName"))) _
        , Trim(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Width"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Height"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Length"))), _
         val(FG.TextMatrix(RowNum, FG.ColIndex("UnitID"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("UnitName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("UnitNamee"))), val(FG.TextMatrix(RowNum, FG.ColIndex("LineID")))
 Else
 addtotable val(FG.TextMatrix(RowNum, FG.ColIndex("PrintedQty"))), FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), val(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))), FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Name")), , , , , FG.TextMatrix(RowNum, FG.ColIndex("LotNO")), FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")), val(FG.TextMatrix(RowNum, FG.ColIndex("Item_ID"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VatYou"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VAT"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Total"))), (FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), TxtSearchCode.text, DBCboClientName.text, val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))), _
        Trim(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("CusName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("ColorName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("SizeName"))) _
        , Trim(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Width"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Height"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("Length"))), _
        val(FG.TextMatrix(RowNum, FG.ColIndex("UnitID"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("UnitName"))), Trim(FG.TextMatrix(RowNum, FG.ColIndex("UnitNamee"))), val(FG.TextMatrix(RowNum, FG.ColIndex("LineID")))

End If
   '            addtotable val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))), FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), val(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))), FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Name")), , , , , FG.TextMatrix(RowNum, FG.ColIndex("LotNO")), FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")), val(FG.TextMatrix(RowNum, FG.ColIndex("Item_ID"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VatYou"))), val(FG.TextMatrix(RowNum, FG.ColIndex("VAT"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Total"))), (FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")))
          
            End If
        End If
        
        

 

    Next RowNum

    printCodes WindowTarget
    'Unload Me
End Sub

Public Sub printCodes(m_PrintTarget As PrintTarget)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim cCompanyInfo As ClsCompanyInfo

    If Dir(App.path & "\Reports\Inventory\" & "BarCode.rpt") = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

   ' MySQL = "SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, "
   ' MySQL = MySQL & "                   dbo.TblPrintBarCode.Color, dbo.TblPrintBarCode.[size], dbo.TblPrintBarCode.class, dbo.TblPrintBarCode.CodeAnalisys, dbo.TblPrintBarCode.ExpiryDate,"
   ' MySQL = MySQL & "                  dbo.TblPrintBarCode.LotNO"
   ' MySQL = MySQL & "  FROM         dbo.TblPrintBarCode LEFT OUTER JOIN"
   ' MySQL = MySQL & "                  dbo.TblItems ON dbo.TblPrintBarCode.item_id = dbo.TblItems.itemid"
 

    MySQL = "  SELECT dbo.TblPrintBarCode.Suppliername,dbo.TblPrintBarCode.Suppliercode , dbo.TblPrintBarCode.Zcode, dbo.TblPrintBarCode.Zcode128 ,code128,  dbo.TblItems.ItemComment  ,  dbo.TblItems.TotalCalories, dbo.TblItems.shortName,   dbo.TblItems.PrintedName,     dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color,"
'MySQL = "  SELECT   dbo.TblItems.* , dbo.TblPrintBarCode.Code, dbo.TblPrintBarCode.PartNo, dbo.TblPrintBarCode.Cost, dbo.TblPrintBarCode.Name, dbo.TblPrintBarCode.Color,"

    MySQL = MySQL & "                      dbo.TblPrintBarCode.size, dbo.TblPrintBarCode.class, dbo.TblPrintBarCode.CodeAnalisys, dbo.TblPrintBarCode.ExpiryDate, dbo.TblPrintBarCode.LotNO, dbo.TblPrintBarCode.VatYou, dbo.TblPrintBarCode.VAT,"
    MySQL = MySQL & "                      dbo.TblPrintBarCode.total,TblPrintBarCode.ProductionDate,"
    
    MySQL = MySQL & "                      dbo.TblPrintBarCode.ColorID,TblPrintBarCode.CusName,"
    MySQL = MySQL & "                      dbo.TblPrintBarCode.ColorName,TblPrintBarCode.SizeName,"
    MySQL = MySQL & "                      dbo.TblPrintBarCode.Remarks,TblPrintBarCode.Width,"
    MySQL = MySQL & "                      dbo.TblPrintBarCode.Height,TblPrintBarCode.Length,"
    
     MySQL = MySQL & "                      dbo.TblPrintBarCode.UnitID,TblPrintBarCode.UnitName,"
     MySQL = MySQL & "                      dbo.TblPrintBarCode.UnitNamee,TblPrintBarCode.LineID"
    


        
        
    MySQL = MySQL & "    FROM            dbo.TblPrintBarCode LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblItems ON dbo.TblPrintBarCode.Item_ID = dbo.TblItems.ItemID"
                         
   RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

     Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & "BarCode.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & "BarCode.rpt", , MySQL

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub

Private Sub LblID_Change()
LodaData
End Sub

 
