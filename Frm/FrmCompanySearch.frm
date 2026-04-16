VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmCompanySearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ „Ê—œ"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "FrmCompanySearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   7905
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtRecordNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   5235
   End
   Begin VB.TextBox TxtCompanyName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3456
      Width           =   5235
   End
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„ »«·ﬂ«„· ›ﬁÿ"
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3876
      Width           =   2385
   End
   Begin VB.TextBox XPTxtComID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2610
      Width           =   5235
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      _cx             =   13944
      _cy             =   4419
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCompanySearch.frx":030A
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
      Height          =   372
      Index           =   0
      Left            =   2016
      TabIndex        =   5
      Top             =   3876
      Width           =   912
      _ExtentX        =   1614
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
      Height          =   372
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Top             =   3876
      Width           =   912
      _ExtentX        =   1614
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
      Height          =   372
      Index           =   2
      Left            =   36
      TabIndex        =   7
      Top             =   3876
      Width           =   912
      _ExtentX        =   1614
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
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”Ã·"
      Height          =   315
      Index           =   2
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label lblSearchtype 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   345
      Index           =   0
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3435
      Width           =   1185
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂÊœ"
      Height          =   315
      Index           =   1
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2610
      Width           =   1185
   End
End
Attribute VB_Name = "FrmCompanySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

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
                FG.rows = 2
                                   If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                Msg = "No Avilable Data"
                End If
                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive
            FG.SetFocus

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 1

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap
    Dim s As String
    Dim rsDummy As ADODB.Recordset
    If Not FG.TextMatrix(FG.Row, 1) = "" Then
    
        If Me.lblSearchtype.Caption = 0 Then
            FrmCompany.Retrive val(FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 1 Then
              FrmBillBuy.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 1000 Then
        FrmReturnpurchases.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
        
        ElseIf Me.lblSearchtype.Caption = 156878 Then
              FrmExpenses4.DCVendor.BoundText = val(FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 2009 Then
            Ageng_all.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 2 Then
              FrmCashing.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
     ElseIf Me.lblSearchtype.Caption = 105 Then
              Order_no_search4.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
    
        ElseIf Me.lblSearchtype.Caption = 3 Then
               FrmPayments.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

     ElseIf Me.lblSearchtype.Caption = 19152 Then
             FrmPayments2.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))


        ElseIf Me.lblSearchtype.Caption = 4 Then
                FrmShowPrice.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
   ElseIf Me.lblSearchtype.Caption = 5 Then
                FrmPO4.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
ElseIf Me.lblSearchtype.Caption = 6 Then
               FrmPO5.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

ElseIf Me.lblSearchtype.Caption = 801 Then
               FrmPO10.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
ElseIf Me.lblSearchtype.Caption = 3005 Then
               FrmExpenses3.DCVendor.BoundText = val(FG.TextMatrix(FG.Row, 1))



ElseIf Me.lblSearchtype.Caption = 8 Then
              FrmPO8.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

ElseIf Me.lblSearchtype.Caption = 20915 Then
                FrmVendorContract.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))


 ElseIf Me.lblSearchtype.Caption = 1122014 Then
                FrmInpout.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

  ElseIf Me.lblSearchtype.Caption = 2122014 Then
        '    RSOwner.DBCboClientName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))

    RSOwner.Retrive val(FG.TextMatrix(FG.Row, 1))

ElseIf Me.lblSearchtype.Caption = 3333 Then
        '    RSOwner.DBCboClientName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))

    FrmOtherCustomers.Retrive val(FG.TextMatrix(FG.Row, 1))


ElseIf Me.lblSearchtype.Caption = 241214 Then
               FrmPO9.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

ElseIf Me.lblSearchtype.Caption = 9 Then
               FrmPayments.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
               
ElseIf Me.lblSearchtype.Caption = 10 Then
              projectsbill.DcbosubContractor.BoundText = val(FG.TextMatrix(FG.Row, 1))
              
ElseIf Me.lblSearchtype.Caption = 6060 Then
              projectsbill_Search.DcbosubContractor.BoundText = val(FG.TextMatrix(FG.Row, 1))

              

ElseIf Me.lblSearchtype.Caption = 1010 Then
        frmSubcontractorContract.DcbosubContractor.BoundText = val(FG.TextMatrix(FG.Row, 1))
        frmSubcontractorContract.Text2.text = Trim(FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
        
        's = "Select FullCode from TblCustemers Where "
        'rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
              
  ElseIf Me.lblSearchtype.Caption = 11 Then
              FrmCitiesDistance.VSFlexGrid1.TextMatrix(FrmCitiesDistance.VSFlexGrid1.Row, FrmCitiesDistance.VSFlexGrid1.ColIndex("CusID")) = val(FG.TextMatrix(FG.Row, 1))
              FrmCitiesDistance.VSFlexGrid1.TextMatrix(FrmCitiesDistance.VSFlexGrid1.Row, FrmCitiesDistance.VSFlexGrid1.ColIndex("Fullcode")) = (FG.TextMatrix(FG.Row, 2))
              FrmCitiesDistance.VSFlexGrid1.TextMatrix(FrmCitiesDistance.VSFlexGrid1.Row, FrmCitiesDistance.VSFlexGrid1.ColIndex("CusName")) = (FG.TextMatrix(FG.Row, 3))
ElseIf Me.lblSearchtype.Caption = 1581 Then
              Ageng_all.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

        ElseIf Me.lblSearchtype.Caption = 1915 Then
             FrmCashing1.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
       ElseIf Me.lblSearchtype.Caption = 9915 Then
                  FrmTypeExchange.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
                         
       ElseIf Me.lblSearchtype.Caption = 2016 Then
                  FrmAttributionContract.dcCustomer.BoundText = val(FG.TextMatrix(FG.Row, 1))
                         
               ElseIf Me.lblSearchtype.Caption = 2020 Then
                  frmReport_Scenes.dcCustomer.BoundText = val(FG.TextMatrix(FG.Row, 1))
                
                      ElseIf Me.lblSearchtype.Caption = 2025 Then
                          FrmStopDealing.dcCustomer1.BoundText = val(FG.TextMatrix(FG.Row, 1))
                        
                        ElseIf Me.lblSearchtype.Caption = 2030 Then
                          FrmStopDealing.dcCustomer.BoundText = val(FG.TextMatrix(FG.Row, 1))
                 ElseIf Me.lblSearchtype.Caption = 100 Then
                         FrmOrderUpload.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
                    ElseIf Me.lblSearchtype.Caption = 101 Then
                         FrmTravelTransactions.DBCboClientName2.BoundText = val(FG.TextMatrix(FG.Row, 1))
                 ElseIf Me.lblSearchtype.Caption = 20160211 Then
                   FrmExchangeRequest.dcCustomer.BoundText = val(FG.TextMatrix(FG.Row, 1))
               ElseIf Me.lblSearchtype.Caption = 20202020 Then
                             FrmAddExceptionDays.dcCustomer.BoundText = val(FG.TextMatrix(FG.Row, 1))
               
        End If

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Cmd_Click (2)
End Sub

Private Sub Form_Activate()
If SystemOptions.UserInterface = ArabicInterface Then
If lblSearchtype.Caption = 2122014 Then
FrmCompanySearch.Caption = "«·»ÕÀ ⁄‰ «·„«·ﬂ"
End If
Else
If lblSearchtype.Caption = 2122014 Then
Me.Caption = "Owner Search..."
End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.Row, FG.ColIndex("Code")) = "" Then
            fg_Click
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

    Dim BG As New ClsBackGroundPic
    Dim StrSQL As String

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    Exit Sub
ErrTrap:

End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("C1")) = IIf(IsNull(rs("C1").value), "", rs("C1").value)
                .TextMatrix(Num, .ColIndex("Code")) = IIf(IsNull(rs("CusID").value), "", val(rs("CusID").value))
                .TextMatrix(Num, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", (rs("Fullcode").value))
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
 Else
 .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
 End If
                .TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rs("Cus_Phone").value), "", Trim(rs("Cus_Phone").value))
                .TextMatrix(Num, .ColIndex("Mobile")) = IIf(IsNull(rs("Cus_mobile").value), "", Trim(rs("Cus_mobile").value))
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
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

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap
    StrSQL = "select * From TblCustemers where type=2"
If Me.lblSearchtype = 9 Or Me.lblSearchtype = 10 Then
StrSQL = "select * From TblCustemers where type=3"
ElseIf Me.lblSearchtype = 2122014 Then
StrSQL = "select * From TblCustemers where type=57"
ElseIf Me.lblSearchtype = 3333 Or Me.lblSearchtype = 1010 Or Me.lblSearchtype = 6060 Then
StrSQL = "select * From TblCustemers where type=3"

End If

    Begin = True

    If XPTxtComID.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and fullcode like N'%" & (XPTxtComID.text) & "%'"
        Else
            StrWhere = StrWhere + " where fullcode like N'%" & (XPTxtComID.text) & "%'"
            Begin = True
        End If
    End If

If SystemOptions.UserInterface = ArabicInterface Then
    If TxtCompanyName.text <> "" Then
        If XPChkSearchType.value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and CusName =N'" & Trim(TxtCompanyName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusName =N'" & Trim(TxtCompanyName.text) & "%'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and CusName LIKE N'%" & Trim(TxtCompanyName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusName LIKE N'%" & Trim(TxtCompanyName.text) & "%'"
                Begin = True
            End If
        End If
    End If
    
    
   '/////////////////
   If TxtRecordNo.text <> "" Then
        If XPChkSearchType.value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and recordno =N'" & Trim(TxtRecordNo.text) & "%'"
            Else
                StrWhere = StrWhere + " where recordno =N'" & Trim(TxtRecordNo.text) & "%'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and recordno LIKE N'%" & Trim(TxtRecordNo.text) & "%'"
            Else
                StrWhere = StrWhere + " where recordno LIKE N'%" & Trim(TxtRecordNo.text) & "%'"
                Begin = True
            End If
        End If
    End If
    '//////////////////
    
    
    
    
Else

If TxtCompanyName.text <> "" Then
        If XPChkSearchType.value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and CusNamee =N'" & Trim(TxtCompanyName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusNamee =N'" & Trim(TxtCompanyName.text) & "%'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and CusNamee LIKE N'%" & Trim(TxtCompanyName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusNamee LIKE N'%" & Trim(TxtCompanyName.text) & "%'"
                Begin = True
            End If
        End If
    End If
    
    
      '/////////////////
   If TxtRecordNo.text <> "" Then
        If XPChkSearchType.value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and recordno =N'" & Trim(TxtRecordNo.text) & "%'"
            Else
                StrWhere = StrWhere + " where recordno =N'" & Trim(TxtRecordNo.text) & "%'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and recordno LIKE N'%" & Trim(TxtRecordNo.text) & "%'"
            Else
                StrWhere = StrWhere + " where recordno LIKE N'%" & Trim(TxtRecordNo.text) & "%'"
                Begin = True
            End If
        End If
    End If
    '//////////////////
    

End If
 
    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub ChangeLang()
    Me.Caption = "Supplier Search..."
    XPLbl(1).Caption = " Code"
    XPLbl(0).Caption = " Name"
    XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
XPLbl(2).Caption = "Record No."
    With Me.FG
        .TextMatrix(0, .ColIndex("Count")) = "Serial"
        .TextMatrix(0, .ColIndex("Fullcode")) = " Code"
        .TextMatrix(0, .ColIndex("Name")) = " Name"
        .TextMatrix(0, .ColIndex("Phone")) = " Phone"
        .TextMatrix(0, .ColIndex("Mobile")) = " Mobile"
    
    End With

End Sub


