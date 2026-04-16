VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMovingSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì«  «· ÕÊÌ·"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11940
   Icon            =   "FrmMovingSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   11940
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
   Begin VB.TextBox TxtCashCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7410
      TabIndex        =   24
      Top             =   5010
      Width           =   2670
   End
   Begin VB.ComboBox CBoBasedON 
      Height          =   315
      ItemData        =   "FrmMovingSearch.frx":038A
      Left            =   210
      List            =   "FrmMovingSearch.frx":038C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2850
      Width           =   1200
   End
   Begin VB.TextBox TxtInspectionReport 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3540
      Width           =   2895
   End
   Begin VB.TextBox XPTxtBillNum 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   7620
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2805
      Width           =   2445
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   75
      TabIndex        =   1
      Top             =   0
      Width           =   11565
      _cx             =   20399
      _cy             =   4842
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
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMovingSearch.frx":038E
      ScrollTrack     =   -1  'True
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
      Editable        =   0
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   2
      Top             =   5430
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   855
      TabIndex        =   3
      Top             =   5430
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   5430
      Width           =   735
      _ExtentX        =   1296
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   5490
      TabIndex        =   9
      Top             =   3930
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboSecondStore 
      Height          =   315
      Left            =   5490
      TabIndex        =   10
      Top             =   4290
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DtDate 
      Height          =   345
      Left            =   8460
      TabIndex        =   11
      Top             =   3195
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   109117441
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker todate 
      Height          =   345
      Left            =   8460
      TabIndex        =   12
      Top             =   3540
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   109117441
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   2250
      TabIndex        =   16
      Top             =   2820
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   5490
      TabIndex        =   17
      Top             =   4650
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboEmp 
      Height          =   315
      Left            =   2250
      TabIndex        =   18
      Top             =   3180
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10350
      TabIndex        =   25
      Top             =   5085
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄ „‰›– «·⁄„·ÌÂ"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5190
      TabIndex        =   23
      Top             =   2820
      Width           =   1230
   End
   Begin VB.Label lblCustomer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„Ì·"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10200
      TabIndex        =   22
      Top             =   4830
      Width           =   1185
   End
   Begin VB.Label SalesPerson 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„‰œÊ»"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5610
      TabIndex        =   21
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»‰«¡ ⁄·Ì"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1530
      TabIndex        =   20
      Top             =   2850
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      Height          =   270
      Index           =   65
      Left            =   5235
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3600
      Width           =   1170
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ï  «—ÌŒ"
      Height          =   315
      Index           =   3
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3660
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Œ“‰ «·„ÕÊ· ≈·ÌÂ"
      Height          =   315
      Index           =   2
      Left            =   10230
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4410
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰  «—ÌŒ"
      Height          =   315
      Index           =   1
      Left            =   10230
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3315
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   315
      Index           =   0
      Left            =   10230
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Œ“‰ «·„ÕÊ· „‰Â"
      Height          =   315
      Index           =   4
      Left            =   10230
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4035
      Width           =   1275
   End
End
Attribute VB_Name = "FrmMovingSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset


Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.searchtype = 15
        FrmCustemerSearch.Show vbModal
    End If
End Sub
Private Sub DBCboClientName_Click(Area As Integer)
'DBCboClientName_Change
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2
                If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                Msg = "No Results"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Retrive

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            DtDate.value = Null
            FG.Rows = 2

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
 
        
        
          If SystemOptions.UserInterface = ArabicInterface Then
                                   Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        
                Else
                Msg = "Error In Criteria"
                End If
                
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim StrQry As String
    Dim RsDetails As ADODB.Recordset

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
    FrmMoving.publicSearch = True
        FrmMoving.Retrive val(FG.TextMatrix(FG.Row, 1))
        
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Cmd_Click (2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.Row, FG.ColIndex("BillNum")) = "" Then
            Fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Set rs = New ADODB.Recordset
    Dim StrList As String
    Dim RsStores As New ADODB.Recordset
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    
 If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT StoreID,StoreName From TblStore"
  Else
     StrSQL = "SELECT StoreID,StoreNamee  From TblStore"
  End If
  
    
    fill_combo Me.DCboStoreName, StrSQL
    fill_combo Me.DCboSecondStore, StrSQL
    

   
    
'    StrSQL = "select * From TblStore"
    RsStores.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT branch_id,branch_name From TblBranchesData"
    Else
        StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
    End If
 
    
    fill_combo dcBranch, StrSQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
  
    With CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "ÿ·» œ«Œ·Ì "
        .AddItem "›« Ê—… „‘ —Ì« "
    End With
  
  If SystemOptions.UserInterface = ArabicInterface Then
   
    StrList = FG.BuildComboList(RsStores, "StoreName", "StoreID")
Else
StrList = FG.BuildComboList(RsStores, "StoreNamee", "StoreID")
End If
    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("ClientNmae")) = "|" & StrList
    End If

    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("StorName")) = "|" & StrList
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
 
 
    Me.Caption = "Search Moving Vouchers"
 
    XPLbl(0).Caption = "VCHR#"
 
    XPLbl(1).Caption = "From"
    XPLbl(2).Caption = "User"
    XPLbl(3).Caption = "To"
    XPLbl(4).Caption = "From Store"
    XPLbl(4).Caption = "To Store"
 
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Delete"
    Cmd(2).Caption = "Exit"
 
    With FG
    
    .TextMatrix(0, .ColIndex("count")) = "I"
        .TextMatrix(0, .ColIndex("Serial")) = "Vchr#"
        .TextMatrix(0, .ColIndex("BillDate")) = " Date"
        .TextMatrix(0, .ColIndex("ClientNmae")) = "From Store "
        .TextMatrix(0, .ColIndex("StorName")) = "To Store"
 
  
    End With

End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("BillNum")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
'                .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
.TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(rs("Noteserial1").value), "", rs("Noteserial1").value)

                .TextMatrix(Num, .ColIndex("BillDate")) = IIf(IsNull(rs("Transaction_Date").value), "", Format((rs("Transaction_Date").value), "yyyy/m/d"))

                If SystemOptions.SysDataBaseType = AccessDataBase Then
                    .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("Transactions.StoreID").value), "", Trim(rs("Transactions.StoreID").value))
                    .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("Transactions_1.StoreID").value), "", Trim(rs("Transactions_1.StoreID").value))
                ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("StoreID").value), "", Trim(rs("StoreID").value))
                    .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("TOStoreID").value), "", Trim(rs("TOStoreID").value))
                End If

            End With

            rs.MoveNext
        Next Num

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql() As String
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String

  '  StrSQL = "Select * From QryMovingItems"

StrSQL = "SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.StoreID, Transactions_1.StoreID AS ToStoreID, "
 StrSQL = StrSQL & "  dbo.Transactions.Transaction_serial , dbo.Transactions.NoteSerial1"
StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
StrSQL = StrSQL & " dbo.TblStore TblStore_1 INNER JOIN"
StrSQL = StrSQL & " dbo.Transactions INNER JOIN"
StrSQL = StrSQL & " dbo.Transactions Transactions_1 ON dbo.Transactions.Transaction_ID = Transactions_1.ReturnID ON TblStore_1.StoreID = Transactions_1.StoreID ON"
StrSQL = StrSQL & " dbo.TblStore.StoreID = dbo.Transactions.StoreID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 10)"
'StrSQL = StrSQL & "  ORDER BY dbo.Transactions.Transaction_ID"
          Begin = True
    If XPTxtBillNum.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.NoteSerial1 like'%" & (XPTxtBillNum.Text) & "%'"
        Else
            StrWhere = StrWhere + " where Transactions.NoteSerial1 like'%" & (XPTxtBillNum.Text) & "%'"
            Begin = True
        End If
    End If

    

    If Not IsNull(DtDate.value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.Transaction_Date >=" & SQLDate(DtDate.value, True) & ""
        Else
            StrWhere = StrWhere + " where Transactions.Transaction_Date >=" & SQLDate(DtDate.value, True) & ""
            Begin = True
        End If
    End If


    If Not IsNull(todate.value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.Transaction_Date <=" & SQLDate(todate.value, True) & ""
        Else
            StrWhere = StrWhere + " where Transactions.Transaction_Date <=" & SQLDate(todate.value, True) & ""
            Begin = True
        End If
    End If
    
    
    If DCboStoreName.BoundText <> "" And DCboStoreName.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.StoreID =" & DCboStoreName.BoundText
        Else
            StrWhere = StrWhere + " where Transactions.StoreID =" & DCboStoreName.BoundText
            Begin = True
        End If
    End If

    
    If DBCboClientName.BoundText <> "" And DBCboClientName.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.CusID =" & DBCboClientName.BoundText
        Else
            StrWhere = StrWhere + " where Transactions.CusID =" & DBCboClientName.BoundText
            Begin = True
        End If
    End If

    If Trim(TxtCashCustomerName.Text) <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.CashCustomerName like'%" & (TxtCashCustomerName.Text) & "%'"
        Else
            StrWhere = StrWhere + " where Transactions.CashCustomerName like'%" & (TxtCashCustomerName.Text) & "%'"
            Begin = True
        End If
    End If

    If Trim(TxtInspectionReport.Text) <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.InspectionReport like'%" & (TxtInspectionReport.Text) & "%'"
        Else
            StrWhere = StrWhere + " where Transactions.InspectionReport like'%" & (TxtInspectionReport.Text) & "%'"
            Begin = True
        End If
    End If


    If dcBranch.BoundText <> "" And dcBranch.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and  Transactions.BranchId =" & dcBranch.BoundText
        Else
            StrWhere = StrWhere + " where Transactions.BranchId=" & dcBranch.BoundText
            Begin = True
        End If
    End If



    If DcboEmp.BoundText <> "" And DcboEmp.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and  Transactions.Emp_ID =" & DcboEmp.BoundText
        Else
            StrWhere = StrWhere + " where Transactions.Emp_ID=" & DcboEmp.BoundText
            Begin = True
        End If
    End If


    If val(CBoBasedON.ListIndex) >= 0 Then
        If Begin = True Then
            StrWhere = StrWhere + " and  Transactions.BillBasedOn =" & val(CBoBasedON.ListIndex)
        Else
            StrWhere = StrWhere + " where Transactions.BillBasedOn=" & val(CBoBasedON.ListIndex)
            Begin = True
        End If
    End If


    If DCboSecondStore.BoundText <> "" And DCboSecondStore.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and  Transactions_1.StoreID =" & DCboSecondStore.BoundText
        Else
            StrWhere = StrWhere + " where Transactions_1.StoreID=" & DCboSecondStore.BoundText
            Begin = True
        End If
    End If


            
    Build_Sql = StrSQL + StrWhere + " order by Transactions.Noteserial1"
    Exit Function
ErrTrap:
End Function

