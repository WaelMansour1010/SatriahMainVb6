VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Order_no_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ  «Ê«„— «·‘—«¡"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9120
   Icon            =   "Order_no_search.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   9120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtStoreID 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   6135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4575
      Width           =   1665
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "»ÕÀ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Order_no_search.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5160
      Width           =   7830
   End
   Begin VB.TextBox TxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   1830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Order_no_search.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ  «Ê«„— «·‘—«¡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   97583105
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   9075
      _cx             =   16007
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Order_no_search.frx":1EF8
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
      Left            =   2040
      TabIndex        =   15
      Top             =   6480
      Width           =   945
      _ExtentX        =   1667
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
      Left            =   1020
      TabIndex        =   16
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
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
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   120
      TabIndex        =   20
      Top             =   4590
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   270
      Index           =   8
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label lblSpecificsearch 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   495
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„Ì· «·„Ê—œ"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "Order_no_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Private Sub BtnFirst_Click()

End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub DBCboClientName_Change()
    TxtCusID.Text = DBCboClientName.BoundText
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
        If Me.RetrunType = 0 Then
            FrmExpenses5.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
        ElseIf Me.RetrunType = 1 Then
            FrmExpenses3.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
        ElseIf Me.RetrunType = 2 Then
            FrmPayments.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
        ElseIf Me.RetrunType = 3 Then
            FrmBillBuy.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
            FrmBillBuy.DcCurrency.Text = Fg.TextMatrix(Fg.Row, 8)
        ElseIf Me.RetrunType = 4 Then

            With FrmExpenses5.Fg_Journal
                .TextMatrix(.Row, .ColIndex("Order_No")) = Fg.TextMatrix(Fg.Row, 1)
            End With
     
        ElseIf Me.RetrunType = 5 Then

            With FrmExpenses3.Fg_Journal
                .TextMatrix(.Row, .ColIndex("Order_No")) = Fg.TextMatrix(Fg.Row, 1)
            End With
    
        ElseIf Me.RetrunType = 6 Then
            FrmProductionOrder.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
     
      ElseIf Me.RetrunType = 61 Then
            FrmProductionOrder.TxtResProductionNo.Text = Fg.TextMatrix(Fg.Row, 1)
           FrmProductionOrder.ProkerId.Text = Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID"))
           
        ElseIf Me.RetrunType = 7 Then
            FrmOutProductionOrder.TxtWorkOrderNO.Text = Fg.TextMatrix(Fg.Row, 1)
     
        ElseIf Me.RetrunType = 8 Then
            frmsalebill.txtorder_no.Text = Fg.TextMatrix(Fg.Row, 1)
     
        ElseIf Me.RetrunType = 9 Then
            FrmInpout.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
     
     
          ElseIf Me.RetrunType = 10 Then
            FrmPO6.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
            
          ElseIf Me.RetrunType = 11 Then
            FrmOut.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
                 
                 ElseIf Me.RetrunType = 12 Then
            FrmShipmentOrder.TxtPONo.Text = Fg.TextMatrix(Fg.Row, 1)
                         ElseIf Me.RetrunType = 14 Then
            FrmPO9.TxtPONo.Text = Fg.TextMatrix(Fg.Row, 1)
                   
                         ElseIf Me.RetrunType = 15 Then
            FrmTypeExchange.TxtOrderNo.Text = Fg.TextMatrix(Fg.Row, 1)
                   FrmTypeExchange.txtTransaction_ID.Text = Fg.TextMatrix(Fg.Row, 9)
                   
                   FrmTypeExchange.DCboCashType121.ListIndex = 1
                  FrmTypeExchange.DBCboClientName.BoundText = Fg.TextMatrix(Fg.Row, Fg.ColIndex("CusID"))
                  
                   
                             ElseIf Me.RetrunType = 16 Then
            FrmProductionPlan.TxtNoteSerial.Text = Fg.TextMatrix(Fg.Row, 1)
                  
               ElseIf Me.RetrunType = 17 Then
             FrmShowPrice.txt_ORDER_NO.Text = Fg.TextMatrix(Fg.Row, 1)
                 ElseIf Me.RetrunType = 18 Then
             FrmPO3.Retrive Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID"))
                   
                 ElseIf Me.RetrunType = 19 Then
             Projects.txtorder_no.Text = Fg.TextMatrix(Fg.Row, 1)
            
                             ElseIf Me.RetrunType = 20 Then
             FrmPO10.TxtPO6.Text = Fg.TextMatrix(Fg.Row, 1) ' Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID"))
   
   
        End If
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    Dim Transaction_Type As Integer
    
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
            Transaction_Type = IIf(IsNull(rs("Transaction_Type").value), "", rs("Transaction_Type").value)
              If Transaction_Type = 21 Then
                .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            Else
                 .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
            End If
            
                .TextMatrix(Num, .ColIndex("remark")) = IIf(IsNull(rs("remark").value), "", Trim(rs("remark").value))
                .TextMatrix(Num, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", Trim(rs("CusID").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If


       If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
                Else
                    .TextMatrix(Num, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", Trim(rs("StoreNamee").value))
                End If
                
                
                .TextMatrix(Num, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("Transaction_Date").value))
                .TextMatrix(Num, .ColIndex("currency_code")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("currency_code").value))
           
                .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
                
            'Transaction_ID
            End With

            rs.MoveNext
        Next Num

  Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
Dcombos.GetStores Me.DCboStoreName
XPDtbBill.value = Date


    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL
    RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap


StrSQL = " SELECT    dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1,dbo.Transactions.Transaction_Type,  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.Transactions.countryid, "
StrSQL = StrSQL & "  dbo.Transactions.order_no, dbo.Transactions.remark, dbo.TblCountriesData.CountryName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "  dbo.Transactions.Closed, dbo.Transactions.Currency_id, dbo.currency.code AS currency_code, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "  dbo.TblStore.StoreNamee , dbo.TblStore.code"
StrSQL = StrSQL & " FROM         dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.currency ON dbo.Transactions.Currency_id = dbo.currency.id LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & " dbo.TblCountriesData ON dbo.Transactions.countryid = dbo.TblCountriesData.CountryID"
    If Me.RetrunType = 3 Then
      '  StrSQL = "Select * From Order_no_details "

        If lblSpecificsearch.Caption = "1" Then
            StrSQL = StrSQL + " Where   ( Transaction_Type=29  )"
        ElseIf lblSpecificsearch.Caption = "2" Then
            StrSQL = StrSQL + " Where   ( Transaction_Type=17)"
        End If

    ElseIf Me.RetrunType = 8 Or Me.RetrunType = 16 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=" & val(lblSpecificsearch.Caption) & ")"

    ElseIf Me.RetrunType = 7 Then
      '  StrSQL = "Select * From Order_no_details "
        StrSQL = StrSQL + " Where  Transaction_Type=26"

    ElseIf Me.RetrunType = 10 Or Me.RetrunType = 12 Or Me.RetrunType = 13 Or Me.RetrunType = 14 Or Me.RetrunType = 60 Or Me.RetrunType = 61 Then
      '  StrSQL = "Select * From Order_no_details "
      
        StrSQL = StrSQL + " Where   ( Transaction_Type=" & val(lblSpecificsearch.Caption) & ")"
     ElseIf Me.RetrunType = 18 Or Me.RetrunType = 19 Then
 
  StrSQL = StrSQL + " Where ( Transaction_Type=6  )"
    Else
      '  StrSQL = "Select * From Order_no_details "
        StrSQL = StrSQL + " Where ( Transaction_Type=6  or Transaction_Type=29 or    Transaction_Type=17)"

    End If

 If Me.RetrunType = 11 Or Me.RetrunType = 12 Then
     
     If (Me.txtorder_no.Text) <> "" Then
        StrSQL = StrSQL + " AND NoteSerial1 like'%" & Me.txtorder_no.Text & "%'"
    End If
    
 Else
 
 
    If (Me.txtorder_no.Text) <> "" Then
        StrSQL = StrSQL + " AND order_no like'%" & Me.txtorder_no.Text & "%'"
    End If

End If

    If DataCombo4.BoundText <> "" Then
        StrWhere = StrWhere + " and countryid =" & DataCombo4.BoundText & " "
    
    End If
    

    If Me.DBCboClientName.BoundText <> "" Then
 
        StrWhere = StrWhere + " and Transactions.CusID =" & Me.DBCboClientName.BoundText & ""
 
    End If

    If Me.DCboStoreName.BoundText <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.storeid  =" & val(Me.DCboStoreName.BoundText)
 
    End If
    
    If Trim(Me.txtremark.Text) <> "" Then
    
        StrWhere = StrWhere + " and remark like '%" & Trim(Me.txtremark.Text) & "%'"
     
    End If

    'StrSQL = StrSQL & "  order by Transaction_ID "
If Me.RetrunType = 15 Then
 StrWhere = StrWhere & " and requestOrOrder=0 "
 End If


    Build_Sql = StrSQL + StrWhere & "  order by dbo.Transactions.Transaction_ID "
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
                Fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Purchase Orders"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label5.Caption = "Country"
    Label4.Caption = "Vendor"
    Label6.Caption = "Remark"
lbl(8).Caption = "Store"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("order_no")) = "order no"
        .TextMatrix(0, .ColIndex("remark")) = "remark  "
        .TextMatrix(0, .ColIndex("CusName")) = "Vendor Name"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = " Date"
        .TextMatrix(0, .ColIndex("CountryName")) = "Country Name"
        .TextMatrix(0, .ColIndex("STORENAME")) = "Store Name"
        
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub

