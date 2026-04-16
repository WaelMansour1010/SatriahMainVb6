VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ArrowsFollowAlarm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ‰»ÌÂ«  «—»«Õ ÊŒ”« ∆— «·«”Â„"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   Icon            =   "ArrowsFollowAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      Height          =   495
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4560
      Width           =   4575
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ”«—…"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "—»Õ"
         Height          =   255
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   11505
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "modflag"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   480
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
               Picture         =   "ArrowsFollowAlarm.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ArrowsFollowAlarm.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰»ÌÂ«  «—»«Õ ÊŒ”« ∆— «·«”Â„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   5880
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
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
      ButtonImage     =   "ArrowsFollowAlarm.frx":245A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   960
      TabIndex        =   8
      Top             =   4680
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
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
      ButtonImage     =   "ArrowsFollowAlarm.frx":27F4
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   3300
      Left            =   0
      TabIndex        =   14
      Top             =   960
      Width           =   11475
      _cx             =   20241
      _cy             =   5821
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
      BackColorBkg    =   16777215
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
      Rows            =   2
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsFollowAlarm.frx":2B8E
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
   Begin MSDataListLib.DataCombo DcboFinMarketId 
      Height          =   315
      Left            =   7920
      TabIndex        =   15
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   600
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   5400
      Width           =   12855
      ExtentX         =   22675
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õœœ «·»Ê—’Â"
      Height          =   255
      Index           =   4
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "ArrowsFollowAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub FillGridWithData()

    '                        .Cell(flexcpBackColor, i, 12, i, 12) = vbRed
  
End Sub

Private Sub BtnCancel_Click()
    Me.Hide
End Sub

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.VSFlexGrid1.PrintGrid " ‰»Ì…  «—»«Õ ÊŒ”«∆— «·«”Â„", True, 2, 1, 1500
End Sub

Private Sub DcboFinMarketId_Change()
    FillCompanyDatagrid (val(Me.DcboFinMarketId.BoundText))
End Sub

Public Sub FillCompanyDatagrid(FinMarketId As Integer)

    If FinMarketId = 0 Then Exit Sub
    'On Error GoTo ErrTrap

    If FinMarketId = 1 Then

        With Me.VSFlexGrid1
            .Rows = 2
            .Clear flexClearScrollable
        End With

        'Frame1.Visible = False
        'NEW_interface = True
        path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwMLj2AXA0_vQGNzY18g18cQKB-JJO8eEGZq4GniE2wUHOBlbOBpREB3cGKRvp9Hfm6qfkFuRDkAgpcLJw!!/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwMFM1MDBJNExWVENMRzMwMjY!/"
        WebBrowser1.Navigate2 path
        Exit Sub
    End If

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Dim BankName As String
    Dim BankID As String
    Dim GroupID As Integer
    Set rs = New ADODB.Recordset
    My_SQL = "select * From ArrowsCompanies where FinMarketId=" & FinMarketId & "  order by CompanySymbol"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim HyperLink As String
    Dim GroupName As String
    Dim CurrencyId As Integer
    Dim CurrencyName As String

    With Me.VSFlexGrid1
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("CompanyId")) = IIf(IsNull(rs.Fields("CompanyId").value), "", rs.Fields("CompanyId").value)
             
                GroupName = ""
                CurrencyName = ""
                .TextMatrix(i, .ColIndex("groupId")) = IIf(IsNull(rs.Fields("groupId").value), "", rs.Fields("groupId").value)
            
                GroupID = val(.TextMatrix(i, .ColIndex("groupId")))
            
                GetArrowsGroupAccount GroupID, , , , , , GroupName
                .TextMatrix(i, .ColIndex("groupName")) = GroupName
            
                .TextMatrix(i, .ColIndex("CurrencyId")) = IIf(IsNull(rs.Fields("CurrencyId").value), "", rs.Fields("CurrencyId").value)
            
                CurrencyId = val(.TextMatrix(i, .ColIndex("CurrencyId")))
            
                GetCurrencyData CurrencyId, , CurrencyName
                .TextMatrix(i, .ColIndex("CurrencyName")) = CurrencyName
            
                .TextMatrix(i, .ColIndex("CompanySymbol")) = IIf(IsNull(rs.Fields("CompanySymbol").value), "", rs.Fields("CompanySymbol").value)
               
                .TextMatrix(i, .ColIndex("CompanyName")) = IIf(IsNull(rs.Fields("CompanyName").value), "", rs.Fields("CompanyName").value)
                       
                .TextMatrix(i, .ColIndex("CurrentValue")) = IIf(IsNull(rs.Fields("CurrentValue").value), 0, rs.Fields("CurrentValue").value)
            
                get_Financial_market_data val(DcboFinMarketId.BoundText), , , HyperLink
                     
                If FinMarketId = 3 Then
                    .TextMatrix(i, .ColIndex("Hyperlink")) = HyperLink
 
                Else
                    .TextMatrix(i, .ColIndex("Hyperlink")) = HyperLink & IIf(IsNull(rs.Fields("CompanySymbol").value), "", rs.Fields("CompanySymbol").value)
                       
                End If
                  
                .TextMatrix(i, .ColIndex("LastPrice")) = "«÷€ÿ Â‰« „— Ì‰"
                 
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Private Sub Form_Load()
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
 
    Dim Dcombos   As New ClsDataCombos
    Dcombos.getFinMarkets DcboFinMarketId

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If

    'FillCompanyDatagrid (Val(Me.DcboFinMarketId.BoundText))
End Sub

Function cahngelang()
    Label1(2).Caption = "Project Invoices Not Payed"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"

    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        '.TextMatrix(0, .ColIndex("id")) = " Bill ID"
        '.TextMatrix(0, .ColIndex("bill_date")) = "Bill Date  "
        '.TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
        '.TextMatrix(0, .ColIndex("End_user_name")) = "End_user_name"
        '.TextMatrix(0, .ColIndex("Sub_user_name")) = "Sub_user_name"
        '.TextMatrix(0, .ColIndex("total")) = "Bill Total"
        '.TextMatrix(0, .ColIndex("ActualTotal")) = "Payed"
        '.TextMatrix(0, .ColIndex("result")) = "Variance"
        '.TextMatrix(0, .ColIndex("resultpercentage")) = "Variance%"

    End With

End Function

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    'On Error GoTo ErrTrap
    'If NEW_interface = False Then Exit Sub
    Dim i As Integer

    Dim objTable As Object
 
    'The ninth table in the page is the Companies List
    Dim startLoad As Integer
    Dim Cols As Integer
    On Error Resume Next
    startLoad = 75
    Dim lastCompanyId As Integer
    lastCompanyId = CStr(new_id("ArrowsCompanies", "CompanyId", "", True))
    Set objTable = WebBrowser1.Document.getElementsByTagName("table").Item(12)

    With Me.VSFlexGrid1
 
        .Rows = objTable.getElementsByTagName("tr").Length - 1
 
        For i = startLoad To .Rows
            Cols = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Length
            Dim HyperLink  As String
            Dim SymbolNo As Integer

            If Cols >= 2 Then
                '      .TextMatrix((i - startLoad) + 1, .ColIndex("LineNo")) = (i - startLoad) + 1
                .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyName")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(0).innerText
      
            End If
      
            Dim CompanyId As Integer
            Dim GroupID As Integer
            Dim CurrencyId As Integer
            Dim currentvalue As Double
            Dim CompanySymbol As String
            Dim GroupName As String
            Dim CurrencyName  As String
      
            If Cols = 14 Then
                HyperLink = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("a")
                SymbolNo = right(HyperLink, 4)
                .TextMatrix((i - startLoad) + 1, .ColIndex("CompanySymbol")) = SymbolNo
       
                .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(1).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("NetLastPrice")) = .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice"))
     
                CompanyId = 0
                GroupID = 0
                CurrencyId = 0
                currentvalue = 0
                CompanySymbol = 0
                CompanySymbol = SymbolNo
                GetArrowsCompanyData CompanyId, CompanySymbol, , GroupID, , currentvalue, CurrencyId

                If CompanyId <> 0 Then
                    .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyId")) = CompanyId
                Else
                    .TextMatrix((i - startLoad) + 1, .ColIndex("CompanyId")) = lastCompanyId
                    lastCompanyId = lastCompanyId + 1
                End If
 
                .TextMatrix((i - startLoad) + 1, .ColIndex("groupid")) = GroupID
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrentValue")) = currentvalue
                Dim profit As Double
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrencyId")) = CurrencyId
 
                If val(currentvalue) > 0 Then
                    profit = val(.TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice"))) - val(.TextMatrix((i - startLoad) + 1, .ColIndex("CurrentValue")))
                    .TextMatrix((i - startLoad) + 1, .ColIndex("Profit")) = Round(profit, 2)
                Else
                    profit = 0
                End If

                If profit > 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbGreen
                ElseIf profit < 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbRed
                ElseIf profit = 0 Then
                    .Cell(flexcpBackColor, (i - startLoad) + 1, 6, (i - startLoad) + 1, 6) = vbWhite
                End If
           
                GroupName = ""
                CurrencyName = ""
                GetArrowsGroupAccount GroupID, , , , , , GroupName
                .TextMatrix((i - startLoad) + 1, .ColIndex("groupName")) = GroupName
           
                GetCurrencyData CurrencyId, , CurrencyName
                .TextMatrix((i - startLoad) + 1, .ColIndex("CurrencyName")) = CurrencyName
 
            End If

        Next i

        '  .AutoSize 0, .Cols - 1, False
        Dim j As Integer
        Dim lastindex As Integer

        For j = .Rows - 1 To 2 Step -1

            If .TextMatrix(j, .ColIndex("CompanyName")) <> "" Then
                lastindex = j + 1
                GoTo LL
            End If

        Next j

LL:
        .Rows = lastindex + 1
    End With

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"

End Sub

