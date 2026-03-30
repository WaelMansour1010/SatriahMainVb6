VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmToolsRepireItemsCost 
   Caption         =   "‘«‘… Ÿ»ÿ „ Ê”ÿ «· ﬂ·›… ··√’‰«›"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   10440
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   8130
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10440
      _cx             =   18415
      _cy             =   14340
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
      _GridInfo       =   $"FrmToolsRepireItemsCost.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   270
         Left            =   15
         TabIndex        =   3
         Top             =   7845
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   840
         Index           =   1
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   10410
         _cx             =   18362
         _cy             =   1482
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
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√’‰«› ·„ Ì Õœœ ·Â« «Ï  ﬂ·›…"
            Height          =   315
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   450
            Width           =   2235
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   30
            Width           =   4185
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   1380
            TabIndex        =   4
            Top             =   30
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   " ÕœÌÀ"
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
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
            ColorButton     =   14871017
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   2730
            TabIndex        =   6
            Top             =   30
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   " Õ„Ì· «·»Ì«‰« "
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ÿ«„ «·„—«œ «· ÕœÌÀ ≈·ÌÂ"
            Height          =   345
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   1725
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6960
         Left            =   15
         TabIndex        =   1
         Top             =   870
         Width           =   10410
         _cx             =   18362
         _cy             =   12277
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmToolsRepireItemsCost.frx":0081
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
Attribute VB_Name = "FrmToolsRepireItemsCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            updatedata

        Case 2
            LoadData
    End Select

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic

    With Me.Fg
        Set GrdBack = New ClsBackGroundPic
        Set .WallPaper = GrdBack.Picture
        .ExtendLastCol = True
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShowAndMove
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.CboType
        .AddItem "‰Ÿ«„ «·„ Ê”ÿ «·„—ÃÕ «·ÃœÌœ"
        .AddItem "‰Ÿ«„ «·”Ì—Ì«· ‰„»— Ê«Œ— ”⁄— ‘—«¡"
    End With

    Me.Width = 12000
    Me.Height = 9500
    Resize_Form Me
End Sub

Private Sub LoadData()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim DblItemValue As Double

    StrSQL = "SELECT TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial," & "dbo.Transactions.Transaction_Date,dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemCode," & "dbo.TblItems.ItemName, dbo.Transaction_Details.ItemSerial,dbo.Transaction_Details.Quantity," & "dbo.Transaction_Details.Price, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.CostTransID," & "dbo.Transaction_Details.ItemDiscountType,dbo.Transaction_Details.ItemDiscount," & "dbo.Transaction_Details.ItemProfit , dbo.Transaction_Details.Id, dbo.TblCustemers.CusName "
    StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN "
    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = " & "dbo.Transaction_Details.Transaction_ID INNER JOIN dbo.TblItems ON dbo.Transaction_Details.Item_ID =" & "dbo.TblItems.ItemID INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID" & " Where(dbo.Transactions.Transaction_Type = 2) "
 
    If Me.Chk.value = vbChecked Then
        StrSQL = StrSQL + "  AND( Transaction_Details.CostTransID IS NULL)"
    End If

    StrSQL = StrSQL + " ORDER BY dbo.Transactions.Transaction_ID, dbo.Transaction_Details.ID "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg
        .Rows = .FixedRows
        .AutoSize 0, .Cols - 1, False

        If rs.BOF Or rs.EOF Then
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If

        .Rows = .FixedRows + rs.RecordCount

        For i = 1 To rs.RecordCount
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
            .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(rs("Item_ID").value), "", rs("Item_ID").value)
            .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            .TextMatrix(i, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
            .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs("Price").value), "", rs("Price").value)
            .TextMatrix(i, .ColIndex("CostPrice")) = IIf(IsNull(rs("CostPrice").value), "", rs("CostPrice").value)
            .TextMatrix(i, .ColIndex("CostTransID")) = IIf(IsNull(rs("CostTransID").value), "", rs("CostTransID").value)
            .TextMatrix(i, .ColIndex("ItemProfit")) = IIf(IsNull(rs("ItemProfit").value), "", rs("ItemProfit").value)
            .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
            .TextMatrix(i, .ColIndex("ItemDiscountType")) = IIf(IsNull(rs("ItemDiscountType").value), 0, rs("ItemDiscountType").value)
            .TextMatrix(i, .ColIndex("ItemDiscount")) = IIf(IsNull(rs("ItemDiscount").value), 0, rs("ItemDiscount").value)
            '---------------------------------------------------------------------
            'Õ”«» ﬁÌ„… «·Œ’„ ⁄·Ï ﬂ· ’‰›
            '⁄‰ ÿ—Ìﬁ ÷—» «·ﬂ„Ì… ›Ï «·”⁄—
            DblItemValue = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price")))
        
            If val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 0 Or val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 1 Then
                '·«ÌÊÃœ Œ’„
                .TextMatrix(i, .ColIndex("DiscountValue")) = 0
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 2 Then
                'Œ’„ ﬁÌ„…
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue - val(.TextMatrix(i, .ColIndex("ItemDiscount")))
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 3 Then
                'Œ’„ ‰”»…
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue * (1 - (val(.TextMatrix(i, .ColIndex("ItemDiscount"))) / 100))
            ElseIf val(.TextMatrix(i, .ColIndex("ItemDiscountType"))) = 4 Then
                'Œ’„ ﬂ«„·(„Ã«‰Ï)·
                .TextMatrix(i, .ColIndex("DiscountValue")) = DblItemValue
            End If

            rs.MoveNext
        Next i
    
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub updatedata()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim Msg As String
    '-------------------------------
    Dim LngItemID As Long
    Dim StrItemSerial As String
    Dim DblItemCost As Double
    Dim StrTransID As String
    Dim DblItemValue As Double

    '------------------------------
    If Me.CboType.ListIndex = -1 Then
        Msg = "»—Ã«¡ ≈Œ Ì«— ÿ—Ìﬁ… Õ”«» «· ﬂ·›…"
        MsgBox Msg, vbExclamation + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
        Exit Sub
    End If

    If Me.CboType.ListIndex = 0 Then
        'Load FrmItemCostShow

        With Me.Fg

            '        Me.ProgBar.Max = .Rows - 1
            For i = 1 To .Rows - 1
        '        FrmItemCostShow.DcboItemName.BoundText = val(.TextMatrix(i, .ColIndex("Item_ID")))
        '        FrmItemCostShow.TxtToInv.text = val(.TextMatrix(i, .ColIndex("Transaction_ID")))
        '        FrmItemCostShow.DoAction
      '          .TextMatrix(i, .ColIndex("NewCostPrice")) = val(FrmItemCostShow.LblLastCost.Caption)
      '          .TextMatrix(i, .ColIndex("NewCostTrans")) = val(FrmItemCostShow.LblCostTransID.Caption)
                '-----------------------------------------------------------------------------------------
                'Õ”«» ﬁÌ„… «·—»Õ «·ÃœÌœ… »⁄œ «· ⁄œÌ·
                'Ì”«ÊÏ ≈Ã„«·Ï ﬁÌ„… «·’‰› „ÿ—ÊÕ „‰Â«
                '{ﬁÌ„… «·Œ’„ + ﬁÌ„… «· ﬂ·›…}
                'ﬁÌ„… «·’‰›
                DblItemValue = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price")))
            
                'ﬁÌ„… «· ﬂ·›…
                DblItemCost = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("NewCostPrice")))
                    
                .TextMatrix(i, .ColIndex("NewItemProfit")) = DblItemValue - (DblItemCost + .TextMatrix(i, .ColIndex("DiscountValue")))
            
                '-----------------------------------------------------------------
                Me.ProgBar.value = i

                DoEvents
            Next i

            '        Me.ProgBar.Max = .Rows - 1
            Me.ProgBar.value = 0
        
            For i = .FixedRows To .Rows - 1
                StrSQL = "Update Transaction_Details"
                StrSQL = StrSQL + " Set Transaction_Details.CostPrice=" & val(.TextMatrix(i, .ColIndex("NewCostPrice")))
                StrSQL = StrSQL + ",ItemProfit=" & val(.TextMatrix(i, .ColIndex("NewItemProfit")))
            
                StrSQL = StrSQL + ",CostTransID=" & val(.TextMatrix(i, .ColIndex("NewCostTrans")))
            
                StrSQL = StrSQL + " Where Transaction_Details.ID=" & val(.TextMatrix(i, .ColIndex("ID")))
                Cn.Execute StrSQL, , adExecuteNoRecords
                Me.ProgBar.value = i

                DoEvents
            Next i
        
        End With

    ElseIf Me.CboType.ListIndex = 1 Then

        With Me.Fg

            '        Me.ProgBar.Max = .Rows - 1
            For i = 1 To .Rows - 1
                LngItemID = val(.TextMatrix(i, .ColIndex("Item_ID")))
                StrItemSerial = Trim(.TextMatrix(i, .ColIndex("ItemSerial")))
                DblItemCost = GetCostItemPrice(LngItemID, 0, StrItemSerial, StrTransID, LastPurPriceType, , , CDate(.TextMatrix(i, .ColIndex("Transaction_Date"))))
            
                .TextMatrix(i, .ColIndex("NewCostPrice")) = DblItemCost
                .TextMatrix(i, .ColIndex("NewCostTrans")) = StrTransID
           
                '----------------------------------------------------------------------
                'Õ”«» ﬁÌ„… «·—»Õ «·ÃœÌœ… »⁄œ «· ⁄œÌ·
                'Ì”«ÊÏ ≈Ã„«·Ï ﬁÌ„… «·’‰› „ÿ—ÊÕ „‰Â«
                '{ﬁÌ„… «·Œ’„ + ﬁÌ„… «· ﬂ·›…}
                'ﬁÌ„… «·’‰›
                DblItemValue = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price")))
            
                'ﬁÌ„… «· ﬂ·›…
                DblItemCost = val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("NewCostPrice")))
                    
                .TextMatrix(i, .ColIndex("NewItemProfit")) = DblItemValue - (DblItemCost + .TextMatrix(i, .ColIndex("DiscountValue")))
            
                StrSQL = "Update Transaction_Details"
                StrSQL = StrSQL + " Set Transaction_Details.CostPrice=" & val(.TextMatrix(i, .ColIndex("NewCostPrice")))
                StrSQL = StrSQL + ",ItemProfit=" & val(.TextMatrix(i, .ColIndex("NewItemProfit")))
            
                StrSQL = StrSQL + ",CostTransID=" & val(.TextMatrix(i, .ColIndex("NewCostTransID")))
            
                StrSQL = StrSQL + " Where Transaction_Details.ID=" & val(.TextMatrix(i, .ColIndex("ID")))
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                Me.ProgBar.value = i
                Me.Caption = i
            
                DoEvents
            Next i

        End With

    End If

End Sub
