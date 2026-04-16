VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmManChooseItems 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÎÊíÇÑ ÃÕäÇÝ ÇáÝÇÊæÑÉ"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   Icon            =   "FrmManChooseItems.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   7770
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   4695
      Left            =   30
      TabIndex        =   1
      Top             =   1320
      Width           =   7695
      _cx             =   13573
      _cy             =   8281
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
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmManChooseItems.frx":0CCA
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7770
      _cx             =   13705
      _cy             =   1773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   192
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "ÈíÇäÇÊ ÇáÝÇÊæÑÉ"
      Align           =   1
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   1
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÇáÝÇÊæÑÉ Ýì ÇáÈÑäÇãÌ:"
         Height          =   255
         Index           =   7
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   570
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÇáÝÇÊæÑÉ Ýì ÇáÈÑäÇãÌ:"
         Height          =   255
         Index           =   6
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÇáÝÇÊæÑÉ Ýì ÇáÈÑäÇãÌ:"
         Height          =   255
         Index           =   5
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÇáÝÇÊæÑÉ Ýì ÇáÈÑäÇãÌ:"
         Height          =   255
         Index           =   4
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÊÇÑíÎ ÇáÝÇÊæÑÉ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2730
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÇÓã ÇáÚãíá:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2730
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÇáÝÇÊæÑÉ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5910
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÑÞã ÝÇÊæÑÉ ÇáãÈíÚÇÊ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5910
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   1785
      End
   End
   Begin ImpulseButton.ISButton ISBXPBtnOK 
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   6180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ãæÇÝÞ"
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
      ButtonImage     =   "FrmManChooseItems.frx":0EC7
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
   Begin ImpulseButton.ISButton ISBXPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÅáÛÇÁ"
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
      ButtonImage     =   "FrmManChooseItems.frx":1261
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÃÕäÇÝ ÇáÝÇÊæÑÉ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   8
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1050
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -60
      X2              =   7785
      Y1              =   6060
      Y2              =   6075
   End
End
Attribute VB_Name = "FrmManChooseItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Public MyForm As Form

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With Me.FG

        Select Case .ColKey(Col)

            Case "Check"
                .Cell(flexcpChecked, .FixedRows, Col, .Rows - 1, Col) = flexUnchecked
                .Cell(flexcpChecked, Row, Col) = flexChecked
        End Select

    End With

End Sub

Private Sub Fg_DblClick()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If Me.MyForm.name = "FrmManCusRecive" Then
            Me.MyForm.RetriveTicketNO val(.TextMatrix(.Row, .ColIndex("TicketNO")))
        Else
            MyForm.TxtTicketNo.text = .TextMatrix(.Row, .ColIndex("TicketNO"))
            MyForm.DCboItemsCode.BoundText = .TextMatrix(.Row, .ColIndex("Item_ID"))
            MyForm.DCboItemsName.BoundText = .TextMatrix(.Row, .ColIndex("Item_ID"))
            MyForm.TxtSerial.text = .TextMatrix(.Row, .ColIndex("ItemSerial"))
            Unload Me
        End If

    End With

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.FG

        Select Case .ColKey(Col)

            Case "Check"

                If .Cell(flexcpChecked, Row, .ColIndex("HaveSerial")) = flexChecked Then
                    Cancel = False
                Else
                    Cancel = True
                End If

            Case Else
                Cancel = True
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    Me.FG.WallPaper = GrdBack.Picture
    Me.FG.Editable = flexEDKbdMouse
    Resize_Form Me

    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub

Private Sub ChangeLang()

    Me.Caption = "Select Items "
    Ele.Caption = "Invoice Data"
 
    lbl(0).Caption = "Bill No"
    lbl(2).Caption = "Cust, Name"
    lbl(3).Caption = "Bill Date"
    lbl(8).Caption = "Items "

    With FG
        .TextMatrix(0, .ColIndex("Serial")) = "I"
        .TextMatrix(0, .ColIndex("TicketNO")) = "Ticket NO"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemSerial")) = "Item Serial"
        .TextMatrix(0, .ColIndex("guaranteeTime")) = "guarantee Time"
        .TextMatrix(0, .ColIndex("EndGuarantee")) = "End Guarantee"
  
    End With

    ISBXPBtnOK.Caption = "OK"
    ISBXPBtnCancel.Caption = "Exit"

End Sub

Public Sub LoadTrans(LngTansID As Long, _
                     IntTransType As GridTransType)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim m_InvDate As Date

    StrSQL = "SELECT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
    StrSQL = StrSQL + "dbo.Transactions.PaymentType, dbo.Transaction_Details.Item_ID," & "dbo.TblItems.ItemCode, dbo.TblItems.ItemName,dbo.TblItems.ItemNamee,dbo.TblItems.HaveSerial, dbo.Transaction_Details.ItemCase," & "dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price," & "dbo.Transaction_Details.guaranteeTime,dbo.TblItems.HaveGuarantee , dbo.TblItems.GuaranteeValue," & "dbo.TblItems.GuaranteeType "
    StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN "
    StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID INNER JOIN"
    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID =" & "dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
    StrSQL = StrSQL + " Where dbo.Transactions.Transaction_ID=" & LngTansID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FG
    
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Me.lbl(4).Caption = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            Me.lbl(5).Caption = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(6).Caption = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            Else
                Me.lbl(6).Caption = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
            End If
      
            m_InvDate = IIf(IsNull(rs("Transaction_Date").value), Date, rs("Transaction_Date").value)
            Me.lbl(7).Caption = DisplayDate(m_InvDate)
        
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("Item_ID").value), "", rs("Item_ID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                Else
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
                End If

                If rs("HaveSerial").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                .TextMatrix(i, .ColIndex("HaveGuarantee")) = IIf(IsNull(rs("HaveGuarantee").value), "", rs("HaveGuarantee").value)
            
                If rs("HaveGuarantee").value = True Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                    .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                    .TextMatrix(i, .ColIndex("HaveGuarantee")) = 1
                    .TextMatrix(i, .ColIndex("guaranteeTime")) = IIf(IsNull(rs("guaranteeTime").value), "", rs("guaranteeTime").value)
                    .TextMatrix(i, .ColIndex("GuaranteeType")) = IIf(IsNull(rs("GuaranteeType").value), 0, rs("GuaranteeType").value)

                    If .TextMatrix(i, .ColIndex("GuaranteeType")) = 0 Then
                        .TextMatrix(i, .ColIndex("EndGuarantee")) = DisplayDate(CDate(DateAdd("m", val(.TextMatrix(i, .ColIndex("guaranteeTime"))), m_InvDate)))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("guaranteeTime")) = .TextMatrix(i, .ColIndex("guaranteeTime")) & " ÔåÑ "
                    
                        Else
                            .TextMatrix(i, .ColIndex("guaranteeTime")) = .TextMatrix(i, .ColIndex("guaranteeTime")) & " Month/s "
                        End If
                    
                    Else
                        .TextMatrix(i, .ColIndex("EndGuarantee")) = DisplayDate(CDate(DateAdd("d", val(.TextMatrix(i, .ColIndex("guaranteeTime"))), m_InvDate)))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("guaranteeTime")) = .TextMatrix(i, .ColIndex("guaranteeTime")) & " Day/s "
                        Else
                
                        End If
                
                    End If

                ElseIf rs("HaveGuarantee").value = False Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80&
                    .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                    .Cell(flexcpFontName, i, 0, i, .Cols - 1) = "Tahoma"
                    .TextMatrix(i, .ColIndex("guaranteeTime")) = ""
                    .TextMatrix(i, .ColIndex("GuaranteeType")) = ""
                    .TextMatrix(i, .ColIndex("HaveGuarantee")) = 0
                End If
            
                '.TextMatrix(I, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set MyForm = Nothing
End Sub

Private Sub ISBXPBtnCancel_Click()
    Unload Me
End Sub

Private Sub ISBXPBtnOK_Click()
    Dim LngX As Long
    Dim Msg As String

    LngX = GetFgCheckCount(Me.FG, FG.ColIndex("Check"))

    If LngX = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "íÌÈ ÅÎÊíÇÑ ÕäÝ æÇÍÏ Úáì ÇáÃÞá"
        Else
            Msg = "Select At Lease One Item"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        If LngX = 1 Then
    
        Else
        End If
    End If

End Sub

Public Sub ShowManStockItems(LngStoreID As Long, _
                             StrStoreName As String)
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If LngStoreID = 0 Then Exit Sub
    Me.Caption = "ÇáÃÕäÇÝ ÇáãæÌæÏÉ Ýì ãÎÒä ÇáÕíÇäÉ"
    Me.Ele.Caption = "ÈíÇäÇÊ ÇáãÎÒä"

    For i = Me.lbl.LBound To Me.lbl.UBound
        Me.lbl(i).Visible = False
    Next i

    Me.lbl(0).Visible = True
    Me.lbl(0).Caption = "ÇÓã ÇáãÎÒä:"
    Me.lbl(4).Visible = True
    Me.lbl(4).Caption = StrStoreName

    With Me.FG

        .Cols = 1
        .Rows = 1
        .FixedRows = 1
        .Rows = .FixedRows + 1
        .FixedCols = 1
        .Cols = 12
        .ColKey(0) = "Serial"
        .ColKey(1) = "Check"
        .ColKey(2) = "TicketNO"
        .ColKey(3) = "Item_ID"
        .ColKey(4) = "ItemCode"
        .ColKey(5) = "ItemName"
        .ColKey(6) = "HaveSerial"
        .ColKey(7) = "ItemSerial"
        .ColKey(8) = "DateGoIN"
        .ColKey(9) = "DateGoOut"
        .ColKey(10) = "CustomerName"
        .ColKey(11) = "MantainceID"
        .TextMatrix(0, .ColIndex("Serial")) = "ãÓáÓá"
        .TextMatrix(0, .ColIndex("Check")) = "ÅÎÊíÇÑ"
        .ColDataType(.ColIndex("Check")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("TicketNO")) = "ÑÞã ÇáÊßÊ"
        .TextMatrix(0, .ColIndex("Item_ID")) = "ÑÞã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ßæÏ ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemName")) = "ÇÓã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "áå ÓíÑíÇá"
        .ColDataType(.ColIndex("HaveSerial")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("ItemSerial")) = "ÇáÓíÑíÇá"
        .TextMatrix(0, .ColIndex("CustomerName")) = "ÇÓã ÇáÚãíá"
        .ColHidden(8) = False
        .ColHidden(9) = True
        .ColHidden(10) = False
        .ColHidden(11) = True
        .TextMatrix(0, .ColIndex("DateGoIN")) = "ÊÇÑíÎ ÇáÏÎæá"
        .ColDataType(.ColIndex("DateGoIN")) = flexDTDate
        .TextMatrix(0, .ColIndex("MantainceID")) = "ÑÞã æÕá ÇáÕíÇäÉ"
    
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignRightCenter
            .FixedAlignment(i) = flexAlignRightCenter
        Next i

        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
        StrSQL = "SELECT QryManStockComplete.QTY, QryManStockComplete.ItemID, QryManStockComplete.ItemCode," & "QryManStockComplete.ItemName,QryManStockComplete.HaveSerial,QryManStockComplete.ItemSerial," & "QryManStockComplete.TicketNO,QryManStockComplete.StoreID,QryManStockComplete.StoreName," & "dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.ReciptNumber, dbo.TblMaintenece.CashCustomerName," & "dbo.TblMaintenece.CusID, dbo.TblCustemers.CusName, dbo.TblMaintenece.DateGoIN "
        StrSQL = StrSQL + " FROM dbo.TblMainteneceDetails INNER JOIN dbo.TblMaintenece ON " & "dbo.TblMainteneceDetails.MaintananceID = dbo.TblMaintenece.MaintananceID INNER JOIN " & " dbo.QryManStockComplete(0) QryManStockComplete ON dbo.TblMainteneceDetails.TicketNO =" & "QryManStockComplete.TicketNO INNER JOIN dbo.TblCustemers ON dbo.TblMaintenece.CusID =" & "dbo.TblCustemers.CusID "
        StrSQL = StrSQL + " WHERE     (dbo.TblMaintenece.ManOperationTypeID = 1)"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)

                If rs("HaveSerial").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                .TextMatrix(i, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)

                If Not (IsNull(rs("DateGoIN").value)) Then
                    .TextMatrix(i, .ColIndex("DateGoIN")) = DisplayDate(rs("DateGoIN").value)
                End If

                .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)

                If Not IsNull(rs("CashCustomerName").value) Then
                    .TextMatrix(i, .ColIndex("CustomerName")) = .TextMatrix(i, .ColIndex("CustomerName")) & "-" & rs("CashCustomerName").value
                End If

                ''            .TextMatrix(I, .ColIndex("MantainceID")) = IIf(IsNull(Rs("MaintananceID").Value), "", Rs("MaintananceID").Value)
                rs.MoveNext
            Next i

        End If

        If .Rows > 1 Then
            .Cell(flexcpFontBold, .FixedRows, .ColIndex("TicketNO"), .Rows - 1, .ColIndex("TicketNO")) = True
            .Cell(flexcpForeColor, .FixedRows, .ColIndex("TicketNO"), .Rows - 1, .ColIndex("TicketNO")) = SysMaronColor
        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Public Sub ShowManSupStock(LngCusID As Long, _
                           StrCusName As String)
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Me.Caption = "ÇáÃÕäÇÝ ÇáãæÌæÏÉ áÏì ÇáãæÑÏ"
    Me.Ele.Caption = "ÈíÇäÇÊ ÇáãæÑÏ"

    For i = Me.lbl.LBound To Me.lbl.UBound
        Me.lbl(i).Visible = False
    Next i

    Me.lbl(0).Visible = True
    Me.lbl(0).Caption = "ÇÓã ÇáãæÑÏ:"
    Me.lbl(4).Visible = True
    Me.lbl(4).Caption = StrCusName

    With Me.FG
        .Cols = 1
        .Rows = 1
        .FixedRows = 1
        .Rows = .FixedRows + 1
        .FixedCols = 1
        .Cols = 8
        .ColKey(0) = "Serial"
        .ColKey(1) = "Check"
        .ColKey(2) = "TicketNO"
        .ColKey(3) = "Item_ID"
        .ColKey(4) = "ItemCode"
        .ColKey(5) = "ItemName"
        .ColKey(6) = "HaveSerial"
        .ColKey(7) = "ItemSerial"
        .TextMatrix(0, .ColIndex("Serial")) = "ãÓáÓá"
        .TextMatrix(0, .ColIndex("Check")) = "ÅÎÊíÇÑ"
        .ColDataType(.ColIndex("Check")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("TicketNO")) = "ÑÞã ÇáÊßíÊ"
        .TextMatrix(0, .ColIndex("Item_ID")) = "ÑÞã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ßæÏ ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemName")) = "ÇÓã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "áå ÓíÑíÇá"
        .ColDataType(.ColIndex("HaveSerial")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("ItemSerial")) = "ÇáÓíÑíÇá"
     
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignRightCenter
            .FixedAlignment(i) = flexAlignRightCenter
        Next i

        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
        StrSQL = "SELECT QryManSupStockComplete.*"
        StrSQL = StrSQL + " FROM dbo.QryManSupStockComplete(0) QryManSupStockComplete"
        StrSQL = StrSQL + " Where QryManSupStockComplete.CusID=" & LngCusID & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)

                If rs("HaveSerial").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                .TextMatrix(i, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
                rs.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Public Sub ShowManTrans(LngManTransID As Long)
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    Me.Caption = "ÇáÃÕäÇÝ ÇáãæÌæÏÉ Ýì ÅíÕÇá ÅÓÊáÇã( ÏÎæá ) ÕíÇäÉ"
    Me.Ele.Caption = "ÈíÇäÇÊ æÕá ÅíÕÇá ÏÎæá ÇáÕíÇäÉ"

    For i = Me.lbl.LBound To Me.lbl.UBound
        Me.lbl(i).Visible = True
        Me.lbl(i).Caption = ""
    Next i

    Me.lbl(0).Visible = True
    Me.lbl(0).Caption = "ÇÓã ÇáãÎÒä:"
    Me.lbl(4).Visible = True
    Me.lbl(4).Caption = StrStoreName

    With Me.FG
        .Cols = 1
        .Rows = 1
        .FixedRows = 1
        .Rows = .FixedRows + 1
        .FixedCols = 1
        .Cols = 8
        .ColKey(0) = "Serial"
        .ColKey(1) = "Check"
        .ColKey(2) = "TicketNO"
        .ColKey(3) = "Item_ID"
        .ColKey(4) = "ItemCode"
        .ColKey(5) = "ItemName"
        .ColKey(6) = "HaveSerial"
        .ColKey(7) = "ItemSerial"
        .TextMatrix(0, .ColIndex("Serial")) = "ãÓáÓá"
        .TextMatrix(0, .ColIndex("Check")) = "ÅÎÊíÇÑ"
        .ColDataType(.ColIndex("Check")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("TicketNO")) = "ÑÞã ÇáÊßÊ"
        .TextMatrix(0, .ColIndex("Item_ID")) = "ÑÞã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ßæÏ ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("ItemName")) = "ÇÓã ÇáÕäÝ"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "áå ÓíÑíÇá"
        .ColDataType(.ColIndex("HaveSerial")) = flexDTBoolean
        .TextMatrix(0, .ColIndex("ItemSerial")) = "ÇáÓíÑíÇá"
      
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignRightCenter
            .FixedAlignment(i) = flexAlignRightCenter
        Next i

        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False
        StrSQL = "SELECT dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.ReciptNumber," & "dbo.TblCustemers.CusID, dbo.TblCustemers.CusName,dbo.TblMaintenece.CashCustomerName," & "dbo.TblMaintenece.DateGoIN, dbo.TblStore.StoreName, dbo.TblMainteneceDetails.ItemID," & "dbo.TblItems.ItemCode, dbo.TblItems.ItemName,dbo.TblItems.HaveSerial,dbo.TblMainteneceDetails.ItemSerial," & "dbo.TblMainteneceDetails.Quantity,dbo.TblMainteneceDetails.TicketNO, " & "dbo.TblMainteneceDetails.CustomerNotes, dbo.TblMainteneceDetails.EmpNotes,"
        StrSQL = StrSQL + " dbo.TblMainteneceDetails.Cost "
        StrSQL = StrSQL + " FROM         dbo.TblMaintenece INNER JOIN "
        StrSQL = StrSQL + " dbo.TblMainteneceDetails ON dbo.TblMaintenece.MaintananceID =" & "dbo.TblMainteneceDetails.MaintananceID INNER JOIN"
        StrSQL = StrSQL + " dbo.TblItems ON dbo.TblMainteneceDetails.ItemID = dbo.TblItems.ItemID INNER JOIN"
        StrSQL = StrSQL + " dbo.TblCustemers ON dbo.TblMaintenece.CusID = dbo.TblCustemers.CusID INNER JOIN"
        StrSQL = StrSQL + " dbo.TblStore ON dbo.TblMaintenece.StoreID = dbo.TblStore.StoreID"
        StrSQL = StrSQL + " WHERE (dbo.TblMaintenece.MaintananceID=" & LngManTransID & ")"
    
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)

                If rs("HaveSerial").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
                .TextMatrix(i, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
                '            .TextMatrix(i, .ColIndex("CustomerName")) = IIf(IsNull(Rs("CusName").Value), "", Rs("CusName").Value)
                '            If Not IsNull(Rs("CashCustomerName").Value) Then
                '                .TextMatrix(i, .ColIndex("CustomerName")) = .TextMatrix(i, .ColIndex("CustomerName")) & _
                '                "-" & Rs("CashCustomerName").Value
                '            End If
                rs.MoveNext
            Next i

        End If

        If .Rows > 1 Then
            .Cell(flexcpFontBold, .FixedRows, .ColIndex("TicketNO"), .Rows - 1, .ColIndex("TicketNO")) = True
            .Cell(flexcpForeColor, .FixedRows, .ColIndex("TicketNO"), .Rows - 1, .ColIndex("TicketNO")) = SysMaronColor
        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

