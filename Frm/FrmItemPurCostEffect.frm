VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmItemPurCostEffect 
   Caption         =   " √ÀÌ— ›« Ê—… ‘—«¡ «Ê —’Ìœ ≈›  «ÕÏ ›Ï √—»«Õ ›Ê« Ì— «·„»Ì⁄« "
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   9945
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9945
      _cx             =   17542
      _cy             =   13176
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmItemPurCostEffect.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   585
         Index           =   1
         Left            =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6870
         Width           =   9915
         _cx             =   17489
         _cy             =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
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
         Begin MSComctlLib.ProgressBar PrgBar 
            Height          =   315
            Left            =   0
            TabIndex        =   4
            Top             =   150
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   975
         Index           =   0
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   9915
         _cx             =   17489
         _cy             =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   405
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   30
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
            Caption         =   " ‰›Ì–"
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
            ColorButton     =   14871017
         End
         Begin VB.TextBox TxtTransID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   360
            Width           =   1905
         End
         Begin VB.ComboBox CboSearchType 
            Height          =   315
            Left            =   6060
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   1845
         End
         Begin VB.ComboBox CboTransType 
            Height          =   315
            Left            =   7950
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   1845
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   405
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   450
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
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
            ColorButton     =   14871017
         End
         Begin VB.Label LBL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄Ì«— «·»ÕÀ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   2
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   90
            Width           =   1785
         End
         Begin VB.Label LBL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   1
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   90
            Width           =   1785
         End
         Begin VB.Label LBL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   0
            Left            =   7980
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   90
            Width           =   1785
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5850
         Left            =   15
         TabIndex        =   1
         Top             =   1005
         Width           =   9915
         _cx             =   17489
         _cy             =   10319
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmItemPurCostEffect.frx":0081
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
         ExplorerBar     =   0
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
   End
End
Attribute VB_Name = "FrmItemPurCostEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Me.LoadTrans

        Case 1
    End Select

End Sub

Private Sub Fg_DblClick()
    Dim Lngid As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub

        Select Case .ColKey(.Col)

            Case "Transaction_ID", "Transaction_Serial", "Transaction_Date"
                Lngid = val(.TextMatrix(.Row, .ColIndex("Transaction_ID")))

                If Lngid <> 0 Then
                    OpenScreen InvoiceScreen, Lngid
                ElseIf .Rowdata(.Row) <> 0 Then
                    OpenScreen PopUpShowItemCardScreen, .Rowdata(.Row), 0
                End If

            Case "CusName"
                Lngid = val(.TextMatrix(.Row, .ColIndex("CusID")))

                If Lngid <> 0 Then
                    OpenScreen PopUpShowCustomerBalanceScreen, Lngid, 0
                End If

            Case "CostPrice"
                Set XNode = .GetNode(.Row)

                If Not XNode Is Nothing Then
                    Lngid = val(XNode.key)

                    If Lngid <> 0 Then
                        'Load FrmItemCostShow
                        'FrmItemCostShow.DcboItemName.BoundText = Lngid
                        'FrmItemCostShow.DoAction
                        'FrmItemCostShow.show
                        'FrmItemCostShow.ZOrder 0
                    End If
                End If

        End Select

    End With

End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    Dim Lngid As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long
    Dim StrToolTip As String

    With Me.Fg
        .ToolTipText = ""

        If .MouseRow = -1 Then Exit Sub
        If .MouseCol = -1 Then Exit Sub
        LngMouseRow = .MouseRow
        LngMouseCol = .MouseCol
    
        Select Case .ColKey(LngMouseCol)

            Case "Transaction_ID", "Transaction_Serial", "Transaction_Date"
                Lngid = val(.TextMatrix(LngMouseRow, .ColIndex("Transaction_ID")))

                If Lngid <> 0 Then
                    StrToolTip = "≈÷€ÿ Â‰« „— Ì ‰ „  «· Ì‰ Õ Ï Ì „ ⁄—÷ ·ﬂ Â–Â «·›« Ê—…"
                ElseIf .Rowdata(.Row) <> 0 Then
                    StrToolTip = "≈÷€ÿ Â‰« „— Ì ‰ „  «· Ì‰ Õ Ï Ì „ ⁄—÷ ·ﬂ ‘«‘…  ﬁ«—Ì— «·’‰›"
                End If

            Case "CusName"
                StrToolTip = "≈÷€ÿ Â‰« „— Ì ‰ „  «· Ì‰ Õ Ï Ì „ ⁄—÷ ·ﬂ  ﬁ—Ì— ﬂ‘› Õ”«» «·⁄„Ì·"

            Case "CostPrice"
                StrToolTip = "≈÷€ÿ Â‰« „— Ì ‰ „  «· Ì‰ Õ Ï Ì „ ⁄—÷ ·ﬂ ‘«‘… Õ—ﬂ…  ﬂ·›… «·’‰›"
        End Select

        .ToolTipText = StrToolTip
    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set Me.Icon = mdifrmmain.ImgLstMenuIcons.ListImages("Execute").ExtractIcon
    Cmd(0).ButtonStyle = impActive
    Cmd(1).ButtonStyle = impActive
    Set Me.Cmd(1).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Print").Picture
    Me.Cmd(1).ButtonPositionImage = impRightOfText

    Set Me.Cmd(0).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Execute").Picture
    Me.Cmd(0).ButtonPositionImage = impRightOfText

    With Me.CboTransType
        .Clear
        .AddItem "›« Ê—… „‘ —Ì« "
        .AddItem "—’Ìœ ≈›  «ÕÏ"
    End With

    With Me.CboSearchType
        .Clear
        .AddItem "—ﬁ„ «·Õ—ﬂ…"
        .AddItem "„”·”· «·Õ—ﬂ…"
    End With

    With Me.Fg
        Set GrdBack = New ClsBackGroundPic
        Set .WallPaper = GrdBack.Picture
        .Rows = .FixedRows
        .RowHeightMin = 300
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .AutoSize 0, .Cols - 1, False
    End With

    Me.Height = 9000
    Me.Width = 12000
    Resize_Form Me
End Sub

Public Sub LoadTrans()
    On Error Resume Next
    Dim RsItems As ADODB.Recordset
    Dim RsTransInvs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSQLInvs As String
    Dim i As Long
    Dim LngItemID As Long
    Dim LngLastItemRow As Long
    Dim Msg As String
    Dim LngTransID As Long

    StrSQL = "SELECT DISTINCT dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode," & "dbo.TblItems.ItemName , dbo.Transaction_Details.Transaction_ID,dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Type "
    StrSQL = StrSQL + " FROM dbo.Transaction_Details INNER JOIN dbo.TblItems ON " & "dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN dbo.Transactions ON " & "dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID "

    If Me.CboTransType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·Õ—ﬂ… ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.CboTransType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Trim(Me.TxtTransID.text) = "" Then
        Msg = "ÌÃ» ≈œŒ«· „⁄Ì«— «·»ÕÀ ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtTransID.SetFocus
        Exit Sub
    End If

    If Me.CboTransType.ListIndex = 0 Then
        StrSQL = StrSQL + " Where dbo.Transactions.Transaction_Type=1 "
    ElseIf Me.CboTransType.ListIndex = 1 Then
        StrSQL = StrSQL + " Where dbo.Transactions.Transaction_Type=3 "
    End If

    If Me.CboSearchType.ListIndex = 0 Then
        LngTransID = val(Me.TxtTransID.text)
        StrSQL = StrSQL + " AND dbo.Transaction_Details.Transaction_ID=" & LngTransID
    ElseIf Me.CboSearchType.ListIndex = 1 Then

        If Me.CboTransType.ListIndex = 0 Then
            LngTransID = GetTransIDSerial(0, , Trim$(Me.TxtTransID.text), 1)
        Else
            LngTransID = GetTransIDSerial(0, , Trim$(Me.TxtTransID.text), 3)
        End If

        StrSQL = StrSQL + " AND dbo.Transaction_Details.Transaction_ID=" & LngTransID
    End If

    Set RsItems = New ADODB.Recordset
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsItems.BOF Or RsItems.EOF Then
        RsItems.Close
        Set RsItems = Nothing
        Exit Sub
    End If

    StrSQLInvs = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type,"
    StrSQLInvs = StrSQLInvs + "dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.Transactions.PaymentType," & "dbo.Transaction_Details.Item_ID,dbo.Transaction_Details.CostPrice,dbo.Transaction_Details.CostTransID," & "dbo.Transaction_Details.Quantity,dbo.Transaction_Details.Price,dbo.Transaction_Details.ItemDiscountType," & "dbo.Transaction_Details.ItemDiscount," & "dbo.Transaction_Details.ItemProfit "
    StrSQLInvs = StrSQLInvs + " FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON " & "dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN "
    StrSQLInvs = StrSQLInvs + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
    StrSQLInvs = StrSQLInvs + " Where dbo.Transaction_Details.CostTransID=" & LngTransID
    StrSQLInvs = StrSQLInvs + " AND dbo.Transaction_Details.Item_ID="

    With Me.Fg
    
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
    
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .GridLines = flexGridNone
        RsItems.MoveFirst

        Do While Not RsItems.EOF
            .AddItem vbTab & RsItems("ItemName").value
            LngLastItemRow = .Rows - 1
            .Rowdata(LngLastItemRow) = RsItems("Item_ID").value
            .IsSubtotal(LngLastItemRow) = True
            .Cell(flexcpFontBold, LngLastItemRow, 1) = True
            .Cell(flexcpForeColor, LngLastItemRow, 1, LngLastItemRow, .Cols - 1) = vbBlue
            '-----------------------------------------------
            LngItemID = RsItems("Item_ID").value
            StrSQL = StrSQLInvs & LngItemID
            StrSQL = StrSQL + " Order By Transactions.Transaction_ID DESC"
            Set RsTransInvs = New ADODB.Recordset
            RsTransInvs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTransInvs.BOF Or RsTransInvs.EOF) Then
                RsTransInvs.MoveFirst
            
                LngLastItemRow = LngLastItemRow + 1

                Do While Not RsTransInvs.EOF
                    .AddItem ""
                    LngLastItemRow = .Rows - 1
                    .TextMatrix(LngLastItemRow, .ColIndex("Transaction_ID")) = RsTransInvs("Transaction_ID").value
                    .TextMatrix(LngLastItemRow, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsTransInvs("Transaction_Serial").value), "", RsTransInvs("Transaction_Serial").value)

                    If Not IsNull(RsTransInvs("Transaction_Date").value) Then
                        .TextMatrix(LngLastItemRow, .ColIndex("Transaction_Date")) = DisplayDate(RsTransInvs("Transaction_Date").value)
                    End If

                    '
                    .TextMatrix(LngLastItemRow, .ColIndex("CusID")) = IIf(IsNull(RsTransInvs("CusID").value), "", RsTransInvs("CusID").value)
                    
                    .TextMatrix(LngLastItemRow, .ColIndex("CusName")) = IIf(IsNull(RsTransInvs("CusName").value), "", RsTransInvs("CusName").value)
                    '«·ﬂ„Ì… «·„»«⁄…
                    .TextMatrix(LngLastItemRow, .ColIndex("Quantity")) = IIf(IsNull(RsTransInvs("Quantity").value), "", RsTransInvs("Quantity").value)
                    '”⁄— «·»Ì⁄
                    .TextMatrix(LngLastItemRow, .ColIndex("Price")) = IIf(IsNull(RsTransInvs("Price").value), "", RsTransInvs("Price").value)
                    '”⁄— «· ﬂ·›…
                    .TextMatrix(LngLastItemRow, .ColIndex("CostPrice")) = IIf(IsNull(RsTransInvs("CostPrice").value), "", RsTransInvs("CostPrice").value)
                 
                    '≈Ã„«·Ï «· ﬂ·›…
                    .TextMatrix(LngLastItemRow, .ColIndex("TotalCost")) = val(.TextMatrix(LngLastItemRow, .ColIndex("Quantity"))) * val(.TextMatrix(LngLastItemRow, .ColIndex("CostPrice")))
                    
                    RsTransInvs.MoveNext
                Loop

            End If

            '-----------------------------------------------
            RsItems.MoveNext
        Loop

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

