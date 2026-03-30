VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmGridAddItemSerials 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· «·”Ì—Ì«·"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11490
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkChooseAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Œ Ì«— «·ﬂ·"
      Height          =   420
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1080
      Width           =   1125
   End
   Begin VB.TextBox txtFile 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TXTPrice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TxtStoreID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9690
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   840
      Width           =   525
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   7
      Top             =   5250
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1350
      Width           =   6345
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   8
      Top             =   5250
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg1 
      Height          =   3495
      Left            =   6600
      TabIndex        =   10
      Top             =   1560
      Width           =   4785
      _cx             =   8440
      _cy             =   6165
      Appearance      =   2
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
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmGridAddItemSerials.frx":0000
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DcboStores 
      Height          =   315
      Left            =   6240
      TabIndex        =   13
      Top             =   840
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton ISButton3 
      Height          =   315
      Left            =   7080
      TabIndex        =   21
      Top             =   5280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   "«” Ì—«œ «·„·›"
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
      ButtonImage     =   "FrmGridAddItemSerials.frx":007D
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton4 
      Height          =   315
      Left            =   9330
      TabIndex        =   22
      ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
      Top             =   5280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   "Õœœ «·„”«—"
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
      ButtonImage     =   "FrmGridAddItemSerials.frx":68DF
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      LowerToggledContent=   0   'False
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6480
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Vatyo 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   495
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Vat 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   495
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblunitname 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblunitid 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”⁄— «·«› —«÷Ì"
      Height          =   255
      Index           =   7
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   315
      Index           =   6
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   870
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label ItemID 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   375
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11500
      X2              =   0
      Y1              =   5160
      Y2              =   5175
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   540
      Width           =   9645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›:"
      Height          =   255
      Index           =   2
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   540
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   8580
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›: "
      Height          =   255
      Index           =   1
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”ÿ—: "
      Height          =   255
      Index           =   0
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "FrmGridAddItemSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fg As VSFlex8UCtl.VSFlexGrid

Public LngRow As Long

Public LngCol As Long

Private Sub chkChooseAll_Click()

 
Dim i As Integer
Dim CurrentStr As String
For i = 1 To FG1.Rows - 1
    If FG1.TextMatrix(i, FG1.ColIndex("Serial")) <> "" Then
            If chkChooseAll.value = 1 Then
                    FG1.TextMatrix(i, FG1.ColIndex("Select")) = 1
                      CurrentStr = FG1.TextMatrix(i, FG1.ColIndex("Serial")) & "," & CurrentStr
                    
                      
            Else
                    FG1.TextMatrix(i, FG1.ColIndex("Select")) = 0
                    Me.TxtComment.Text = ""
            End If
    End If
Next

      If CurrentStr <> "" Then
            Me.TxtComment.Text = mId(CurrentStr, 1, Len(CurrentStr) - 1)
        End If
        
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String





 
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
  '  On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    'strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(TxtComment, strFilterText)
    Dim i As Integer
 
   
    'Num = currentrow

     
    Dim CurrentSerial As String
          
'******************************************************'
For intX = 0 To UBound(astrSplitItems)
   CurrentSerial = astrSplitItems(intX)
   Dim intj As Integer
       For intj = 0 To UBound(astrSplitItems)
      If CurrentSerial = astrSplitItems(intj) Then
                    If intj = intX Then
                    
                    Else
           MsgBox "«·”Ì—Ì«· „ﬂ—— " & CurrentSerial
               GoTo ErrTrap
                    End If
     
      
      End If
      
       
       Next intj
   
 Next
 
    If Not fg Is Nothing Then
    
        'FG.TextMatrix(LngRow, LngCol) = Trim$(Me.TxtComment.text)
  
  
  
        
         mdifrmmain.ActiveForm.RetriveSerials ItemID.Caption, lbl(5).Caption, Trim$(Me.TxtComment.Text), LngRow, val(Me.txtPrice.Text), lblunitid.Caption, Me.lblunitname
        
        'frmsalebill.RetriveSerials ItemID.Caption, lbl(5).Caption, Trim$(Me.TxtComment.text), LngRow
 
        Unload Me
    End If
ErrTrap:
End Sub

Private Sub Fg1_AfterEdit(ByVal Row As Long, _
                          ByVal Col As Long)

    Dim IntCounter As Integer
    Dim CurrentStr As String
    Dim i As Integer
    CurrentStr = ""

    With Me.FG1

        For i = .FixedRows To .Rows - 1
     
            If .TextMatrix(i, .ColIndex("Serial")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
     
                CurrentStr = .TextMatrix(i, .ColIndex("Serial")) & "," & CurrentStr
  
            End If

        Next i

        If CurrentStr <> "" Then
            Me.TxtComment.Text = mId(CurrentStr, 1, Len(CurrentStr) - 1)
        End If

    End With

End Sub

Private Sub Fg1_BeforeEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim Msg As String

    With FG1
 
        Select Case .ColKey(Col)

            Case "Select"
                Cancel = False
                Exit Sub
        
        End Select
                 
        Cancel = True
    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores Me.DcboStores
    'Set cDcboSearch(1) = New clsDCboSearch
    'Set cDcboSearch(1).Client = Me.DcboStores
 

    CenterForm Me

    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.cmdCancel.ButtonStyle = impActive
    Set cmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    cmdCancel.ButtonPositionImage = impRightOfText

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Public Sub FillSerials(Item_ID As Long, _
                       StoreId As Long, _
                       Transaction_ID As Long)

    Dim StrSQL  As String
    StrSQL = "SELECT     dbo.Transaction_Details.ItemSerial, SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS actqty"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "   dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
    StrSQL = StrSQL & "  WHERE     (dbo.TransactionTypes.StockEffect <> 0) AND (dbo.Transactions.StoreID = " & StoreId & ") AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND"
    StrSQL = StrSQL & "  (dbo.Transaction_Details.Transaction_ID <> " & Transaction_ID & ")"
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.ItemSerial"
    StrSQL = StrSQL & "   Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
 
    FG1.Clear flexClearScrollable, flexClearEverything

    Dim RsUnitData As New ADODB.Recordset
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim Num As Integer
 
    If Not (RsUnitData.EOF Or RsUnitData.BOF) Then
    
        FG1.Rows = RsUnitData.RecordCount + 1

        For Num = 1 To RsUnitData.RecordCount
        
            With FG1
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(RsUnitData("ItemSerial").value), "", (RsUnitData("ItemSerial").value))
            End With
        
            RsUnitData.MoveNext
        
        Next Num

        '    FG1.AutoSize 0, FG1.Cols - 1, False
    End If

    RsUnitData.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ChangeLang()

    Me.Caption = "Items Serials"
lbl(7).Caption = "Price"
    lbl(1).Caption = "Code"
    Me.lbl(2).Caption = "Name"
    CmdOk.Caption = "Save"
    cmdCancel.Caption = "Close"
    lbl(6).Caption = "Store"

    With Me.FG1
        .TextMatrix(0, .ColIndex("NumIndex")) = "Index"
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("serial")) = "Serial"

    End With

End Sub

Private Sub ISButton3_Click()

    On Error Resume Next

    Dim astrSplit2tems2() As String
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long
    Dim VATNO As String
    Dim SerialNo As String
    Dim DocDate As String
    Dim Branch As String
    Dim BranchID As Double
    Dim CusID As Double
    Dim cus As String
    Dim value As String
    Dim VATPer As String
    Dim Notes As String
    Dim TypeService As String
    Dim store As String
    Dim StoreId As Double
    Dim PayedTyp As Integer
    Dim Msg As String
    Dim PaymentNam As String
    Dim Account_Nam As String
    Dim BoxNam As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    If 1 = 1 Then
  '      Me.VSFlexGrid.Clear flexClearScrollable, flexClearEverything
  '      VSFlexGrid.Rows = 1
        If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
        Set ExcelObj = CreateObject("Excel.Application")
        Set ExcelSheet = CreateObject("Excel.Sheet")
        ExcelObj.Workbooks.Open txtFile.Text
        DoEvents
        Set ExcelBook = ExcelObj.Workbooks(1)
        Set ExcelSheet = ExcelBook.Worksheets(1)
        Dim addserial As String
        With ExcelSheet
            i = 1
            Do Until .cells(i, 1) & "" = ""
            SerialNo = .cells(i, 1)
             addserial = addserial & SerialNo & ","
             
            If .cells(i, 1) & "" = "" Then Exit Sub
                i = i + 1
                Loop
        End With
        'ReLineGrid
   '     Me.VSFlexGrid.SetFocus
        ExcelObj.Workbooks.Close

        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    End If
If Len(addserial) > 1 Then
addserial = mId(addserial, 1, Len(addserial) - 1)
End If
TxtComment.Text = addserial
End Sub

Private Sub ISButton4_Click()
'    VSFlexGrid.Clear flexClearScrollable, flexClearEverything
'    VSFlexGrid.Rows = 1
TxtComment.Text = ""
    CD1.ShowOpen
    txtFile.Text = CD1.filename

End Sub
