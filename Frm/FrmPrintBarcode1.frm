VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPrintBarcode1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… «·»«—ﬂÊœ"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "FrmPrintBarcode1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   10800
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
      Picture         =   "FrmPrintBarcode1.frx":038A
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
      Top             =   990
      Width           =   10770
      _cx             =   18997
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmPrintBarcode1.frx":0714
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
      ButtonImage     =   "FrmPrintBarcode1.frx":0834
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
      ButtonImage     =   "FrmPrintBarcode1.frx":0BCE
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
      ButtonImage     =   "FrmPrintBarcode1.frx":0F68
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
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
      Width           =   10725
   End
End
Attribute VB_Name = "FrmPrintBarcode1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Print")) = False
            Next i

        End With

    End If

 

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim RowNum As Integer
    Dim ItemCount As Integer
    'cBarcode.AddItem
    cBarcode.ClearItems

    For RowNum = 1 To FG.Rows - 1

        If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("Print")) = flexChecked Then
            If Not IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))) Then

                For ItemCount = 1 To val(FG.TextMatrix(RowNum, FG.ColIndex("Qty")))
                    cBarcode.AddItem FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), FG.TextMatrix(RowNum, FG.ColIndex("Name")) & "/" & FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Cost"))
                Next ItemCount

            End If
        End If

    Next RowNum

    'cBarcode.AddItem FG.TextMatrix(2, FG.ColIndex("Code")), "yasser", "ahmed"
    FrmSetting.show vbModal

End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col = FG.ColIndex("Name") Or Col = FG.ColIndex("Code") Then
        Cancel = True
    End If

End Sub

Private Sub Form_Activate()
    LodaData
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As New ClsBackGroundPic
  
    CenterForm Me

    FormPostion Me, GetPostion
    LoadIcon

    Check17.value = vbChecked
  Set FG.WallPaper = BGround.Picture
  
        Exit Sub
        
ErrTrap:
End Sub

Private Sub LoadIcon()
    On Error GoTo ErrTrap

    With FG
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Code")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Cost")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Qty")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Print")) = mdifrmmain.ImgLstTree.ListImages("Print").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
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

    If LblID.Caption = "" Then Exit Sub
  '  StrSQL = "SELECT * FROM QryBarcode WHERE Transaction_ID=" & val(LblID.Caption)
  'StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.Quantity, "
  'StrSQL = StrSQL & "   dbo.TblItems.SallingPrice , dbo.TblItems.PartNo, dbo.TblItems.ItemNamee, dbo.TblItems.barCodeNO"
'StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
'StrSQL = StrSQL & "  dbo.TblItems INNER JOIN"
'StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'StrSQL = StrSQL & "  WHERE Transactions.Transaction_ID=" & val(lblid.Caption)
    
    
 StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.Transaction_Details.Quantity, "
StrSQL = StrSQL & " dbo.TblItems.PartNo, dbo.TblItems.ItemNamee, dbo.TblItems.barCodeNO, dbo.TblItemsUnits.UnitSalesPrice AS SallingPrice, dbo.Transaction_Details.UnitId"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & " dbo.TblItems INNER JOIN"
StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & " dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
StrSQL = StrSQL & " dbo.TblItemsUnits ON dbo.Transaction_Details.UnitId = dbo.TblItemsUnits.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID"
StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_ID = " & val(LblID.Caption) & ")"

    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (rs.EOF Or rs.BOF) Then

        With FG
            .Rows = rs.RecordCount + 1

            For RowNum = 1 To rs.RecordCount
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
              Else
                 .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
              End If
              
                    .TextMatrix(RowNum, .ColIndex("barcodeno")) = IIf(IsNull(rs("barcodeno").value), "", rs("barcodeno").value)
                    
             
             
                .TextMatrix(RowNum, .ColIndex("Code")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(RowNum, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
            
                .TextMatrix(RowNum, .ColIndex("Cost")) = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                .TextMatrix(RowNum, .ColIndex("Qty")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
                .Cell(flexcpChecked, RowNum, .ColIndex("Print")) = flexChecked
                rs.MoveNext
            Next RowNum

        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Function addtotable(NoOfRow As Integer, code As String, cost As Double, Optional PartNo As String = "", Optional name As String = "" _
, Optional NameE As String, Optional Color As String, Optional size As String, Optional class As String)
    Dim rs As New ADODB.Recordset
    Dim str As String
    Dim i As Integer

    str = "select * from   TblPrintBarCode where 1=-1"
    
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  For i = 1 To NoOfRow
        rs.AddNew
        rs("PartNo").value = PartNo
        rs("code").value = code
        rs("cost").value = val(cost)
        rs("Name").value = name
        rs("NameE").value = NameE
        rs("Color").value = Color
        rs("size").value = size
        rs("class").value = class
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
    For RowNum = 1 To FG.Rows - 1

        If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("Print")) = flexChecked Then
            If Not IsNull(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))) Then
           
                addtotable val(FG.TextMatrix(RowNum, FG.ColIndex("Qty"))), FG.TextMatrix(RowNum, FG.ColIndex("barcodeno")), val(FG.TextMatrix(RowNum, FG.ColIndex("Cost"))), FG.TextMatrix(RowNum, FG.ColIndex("PartNo")), FG.TextMatrix(RowNum, FG.ColIndex("Name"))
          
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

    MySQL = "Select * From TblPrintBarCode "

    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = EnglishInterface Then
  
    Else
       
        Set xReport = xApp.OpenReport(App.path & "\Reports\Inventory\" & "BarCode.rpt")
        xReport.Database.SetDataSource RsData
        Set cCompanyInfo = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabComment
        
    End If

    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title

    Set CViewer = New ClsReportViewer
hide_logo = True
    CViewer.FireReport xReport, m_PrintTarget, "", , , 790, App.path & "\Reports\Inventory\" & "BarCode.rpt"

    Set xApp = Nothing
    Set xReport = Nothing
    Screen.MousePointer = vbDefault
    hide_logo = False
End Sub

