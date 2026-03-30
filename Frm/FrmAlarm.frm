VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAlarm 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊäÈíå"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "FrmAlarm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   5415
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   3135
      Left            =   15
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      _cx             =   9551
      _cy             =   5530
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
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmAlarm.frx":038A
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
      ExplorerBar     =   2
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
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4590
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÎÑæÌ"
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
      ButtonImage     =   "FrmAlarm.frx":04CB
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
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇáÊí ÊãÊ Úáì ÇáÃÕäÇÝ ÇáãÓÌáÉ Ýí åÐå ÇáÝÇÊæÑÉ"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   3
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4170
      Width           =   5115
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÅÐÇ ßäÊ ÊÑÛÈ Ýí ÅÊãÇã åÐå ÇáÚãáíÉ  íÌÈ ÍÐÝ ÇáÚãáíÇÊ "
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   83
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3870
      Width           =   5115
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "áÞÏ Êã ÅÌÑÇÁ ÈÚÖ ÇáÚãáíÇÊ Úáì ÈÚÖ ÇáÃÕäÇÝ ÇáÊí Êã ÔÑÇÆåÇ Ýí åÐå ÇáÝÇÊæÑÉ "
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   105
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   330
      Width           =   5235
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "áÇ íãßä ÅÊãÇã åÐå ÇáÚãáíÉ"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   908
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "FrmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DealingForm As GridTransType

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Retrive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdExit_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    CenterForm Me

    FormPostion Me, GetPostion
    'FG.ColComboList(FG.ColIndex("TransType")) = "#2;ÈíÚ|#5;ãÑÊÌÚ ãÔÊÑíÇÊ"
    Dim BGround As New ClsBackGroundPic
    FG.WallPaper = BGround.Picture
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub Retrive()
    On Error GoTo ErrTrap
    Dim rs As ADODB.Recordset
    Dim RowNum As Integer
    Dim count As Integer
    Dim StrSQL As String
    Dim Num As Integer
    Num = 1

    Select Case Me.DealingForm

        Case PurchaseTransaction

            For RowNum = 1 To FrmBillBuy.FG.Rows - 1

                With FrmBillBuy.FG
                    StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(FrmBillBuy.XPDtbBill.value, True) & ""
                    StrSQL = StrSQL + " and Item_ID=" & .TextMatrix(RowNum, .ColIndex("Code"))

                    If .TextMatrix(RowNum, .ColIndex("HaveSerial")) <> "" Then
                        If .TextMatrix(RowNum, .ColIndex("HaveSerial")) = True Then
                            StrSQL = StrSQL + " and ItemSerial='" & .TextMatrix(RowNum, .ColIndex("Serial")) & "'"
                        End If
                    End If

                End With

                StrSQL = StrSQL + " order by Transaction_ID"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.EOF Or rs.BOF) Then
                    FG.Rows = FG.Rows + rs.RecordCount

                    For count = 1 To rs.RecordCount
                        FG.TextMatrix(Num, FG.ColIndex("Index")) = Num
                        FG.TextMatrix(Num, FG.ColIndex("BillNum")) = IIf(rs("Transaction_ID").value = "", Null, rs("Transaction_ID").value)
                        FG.TextMatrix(Num, FG.ColIndex("Transaction_Serial")) = IIf(rs("Transaction_Serial").value = "", Null, rs("Transaction_Serial").value)
                        FG.TextMatrix(Num, FG.ColIndex("Date")) = IIf(rs("Transaction_Date").value = "", Null, Format(rs("Transaction_Date").value, "YYYY/M/D"))
                        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(rs("ItemName").value = "", Null, rs("ItemName").value)
                   
                        FG.TextMatrix(Num, FG.ColIndex("TransType")) = IIf(rs("TransactionTypeName").value = "", Null, rs("TransactionTypeName").value)
                   
                        FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(rs("Quantity").value = "", Null, rs("Quantity").value)
                        Num = Num + 1
                        rs.MoveNext
                    Next count

                End If

                rs.Close
            Next RowNum

        Case OpeningBalance

            For RowNum = 1 To FrmOpeningBalance.FG.Rows - 1

                With FrmOpeningBalance.FG
                    StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(FrmOpeningBalance.XPDtbBill.value, True) & ""
                    StrSQL = StrSQL + " and Item_ID=" & .TextMatrix(RowNum, .ColIndex("Code"))

                    If .TextMatrix(RowNum, .ColIndex("HaveSerial")) <> "" Then
                        If .TextMatrix(RowNum, .ColIndex("HaveSerial")) = True Then
                            StrSQL = StrSQL + " and ItemSerial='" & .TextMatrix(RowNum, .ColIndex("Serial")) & "'"
                        End If
                    End If

                End With

                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.EOF Or rs.BOF) Then
                    FG.Rows = FG.Rows + rs.RecordCount

                    For count = 1 To rs.RecordCount
                        FG.TextMatrix(Num, FG.ColIndex("Index")) = Num
                        FG.TextMatrix(Num, FG.ColIndex("BillNum")) = IIf(rs("Transaction_ID").value = "", Null, rs("Transaction_ID").value)
                        FG.TextMatrix(Num, FG.ColIndex("Date")) = IIf(rs("Transaction_Date").value = "", Null, Format(rs("Transaction_Date").value, "YYYY/M/D"))
                        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(rs("ItemName").value = "", Null, rs("ItemName").value)
                        FG.TextMatrix(Num, FG.ColIndex("TransType")) = IIf(rs("TransactionTypeName").value = "", Null, rs("TransactionTypeName").value)
                        FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(rs("Quantity").value = "", Null, rs("Quantity").value)
                        Num = Num + 1
                        rs.MoveNext
                    Next count

                End If

                rs.Close
            Next RowNum

        Case InvoiceTransaction

            For RowNum = 1 To frmsalebill.FG.Rows - 1

                With frmsalebill.FG
                    StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(frmsalebill.XPDtbBill.value, True) & ""
                    StrSQL = StrSQL + " and Item_ID=" & .TextMatrix(RowNum, .ColIndex("Code"))
                    StrSQL = StrSQL + " and Transaction_Type=9"

                    If .TextMatrix(RowNum, .ColIndex("HaveSerial")) <> "" Then
                        If .TextMatrix(RowNum, .ColIndex("HaveSerial")) = True Then
                            StrSQL = StrSQL + " and ItemSerial='" & .TextMatrix(RowNum, .ColIndex("Serial")) & "'"
                        End If
                    End If

                End With

                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.EOF Or rs.BOF) Then
                    FG.Rows = FG.Rows + rs.RecordCount

                    For count = 1 To rs.RecordCount
                        FG.TextMatrix(Num, FG.ColIndex("Index")) = Num
                        FG.TextMatrix(Num, FG.ColIndex("BillNum")) = IIf(rs("Transaction_ID").value = "", Null, rs("Transaction_ID").value)
                        FG.TextMatrix(Num, FG.ColIndex("Date")) = IIf(rs("Transaction_Date").value = "", Null, Format(rs("Transaction_Date").value, "YYYY/M/D"))
                        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(rs("ItemName").value = "", Null, rs("ItemName").value)
                        FG.TextMatrix(Num, FG.ColIndex("TransType")) = IIf(rs("TransactionTypeName").value = "", Null, rs("TransactionTypeName").value)
                        FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(rs("Quantity").value = "", Null, rs("Quantity").value)
                        Num = Num + 1
                        rs.MoveNext
                    Next count

                End If

                rs.Close
            Next RowNum

    End Select

    Exit Sub
ErrTrap:
End Sub

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)

    If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
        m_DealingForm = vNewValue
    End If

End Property
