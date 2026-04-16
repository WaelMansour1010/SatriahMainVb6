VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFollowAdminMeasuresearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰  ⁄ﬁÌ» ‘»√‰ ≈Ã—«¡ ≈œ«—Ì"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "FmFollowUpAdminMeasureSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Index           =   2
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3120
      Width           =   3795
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   2055
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94175235
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94175235
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   660
         Width           =   255
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1005
      _ExtentX        =   1773
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
      Left            =   3720
      TabIndex        =   1
      Top             =   3840
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
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo dcdep 
      Height          =   315
      Left            =   9720
      TabIndex        =   3
      Top             =   1200
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcjop 
      Height          =   315
      Left            =   9840
      TabIndex        =   5
      Top             =   480
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboEmployeesName 
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   2760
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6525
      _cx             =   11509
      _cy             =   4630
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
      FormatString    =   $"FmFollowUpAdminMeasureSearch.frx":030A
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
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   135
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   285
      Index           =   9
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÊŸÌ›…"
      Height          =   285
      Index           =   8
      Left            =   2490
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÊŸÌ›…"
      Height          =   285
      Index           =   4
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Êﬁ⁄"
      Height          =   285
      Index           =   2
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1035
   End
End
Attribute VB_Name = "FrmFollowAdminMeasuresearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch






Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            GetData

        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
             '   Me.lbl(0).Caption = "Search Results"
            End If

       Case 2
            Unload Me
    End Select

End Sub



Private Sub fg_Click()

      With Me.fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
       ' If val(.TextMatrix(.Row, .ColIndex("id"))) = 0 Then
       '     Exit Sub
       ' End If

       
                FrmFolloAdminMeasure.Retrive val(.TextMatrix(.Row, .ColIndex("id")))
    
'Unload Me
    End With

End Sub







Private Sub Fg_EnterCell()
With Me.fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("id"))) = 0 Then
            Exit Sub
        End If

            FrmFolloAdminMeasure.Retrive val(.TextMatrix(.Row, .ColIndex("id")))
          

    End With
End Sub



Private Sub Form_Activate()
PutFormOnTop Me.hwnd
End Sub

Private Sub Form_Load()
    
  Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetempyeesFollow Me.DCboEmployeesName
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DCboEmployeesName
    Dcombos.GetEmpLocations Dcdep
    Dcombos.GetEmpJobsTypes dcjop
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    
    
       


End Sub
Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = "SELECT     dbo.TblFollowUpAdminMeasure.RecordDate, dbo.TblBranchesData.branch_name, dbo.TblFollowUpAdminMeasure.ID, dbo.TblBranchesData.branch_namee,"
  StrSQL = StrSQL & "                    dbo.TblFollowUpAdminMeasure.FromName , dbo.TblFollowUpAdminMeasure.PersonalDept"
StrSQL = StrSQL & " FROM         dbo.TblFollowUpAdminMeasure INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblFollowUpAdminMeasure.BranchID = dbo.TblBranchesData.branch_id"
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasure.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblFollowUpAdminMeasure.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasure.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblFollowUpAdminMeasure.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If

    If Me.DCboEmployeesName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasure.ID=" & Me.DCboEmployeesName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblFollowUpAdminMeasure.ID=" & Me.DCboEmployeesName.BoundText & ""
        End If
    End If
       'If Me.dcdep.BoundText <> "" Then
       ' If BolBegine = True Then
       '     StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasureDetails.ProjectName='" & Me.dcdep.text & "'"
       ' Else
       '     BolBegine = True
       '     StrWhere = " Where dbo.TblFollowUpAdminMeasureDetails.ProjectName='" & Me.dcdep.text & "'"
       ' End If
    'End If
    '   If Me.dcjop.BoundText <> "" Then
    '    If BolBegine = True Then
    '        StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasureDetails.JobName='" & Me.dcjop.text & "'"
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblFollowUpAdminMeasureDetails.JobName='" & Me.dcjop.text & "'"
    '    End If
    'End If

   ' If Me.DCUser.BoundText <> "" Then
   '     If BolBegine = True Then
   '         StrWhere = StrWhere & " AND dbo.TblEmpAdvance.UserID=" & Me.DCUser.BoundText & ""
   '     Else
   '         BolBegine = True
   '         StrWhere = " Where dbo.TblEmpAdvance.UserID=" & Me.DCUser.BoundText & ""
   '     End If
   ' End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasure.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblFollowUpAdminMeasure.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblFollowUpAdminMeasure.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblFollowUpAdminMeasure.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
   StrSQL = StrSQL & " Order By dbo.TblFollowUpAdminMeasure.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
              '  Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
              ' .TextMatrix(i, .ColIndex("ProfCode")) = val(IIf(IsNull(rs("Code").value), "", rs("Code").value))
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("AdvanceDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
           ' .TextMatrix(i, .ColIndex("procedure")) = IIf(IsNull(rs("Procedures").value), "", rs("Procedures").value)
                .TextMatrix(i, .ColIndex("ProfNme")) = IIf(IsNull(rs("FromName").value), "", rs("FromName").value)
                .TextMatrix(i, .ColIndex("id")) = val(IIf(IsNull(rs("ID").value), "", rs("ID").value))
           '     .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobName").value), "", rs("JobName").value)
           '     .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("ProjectName").value), "", rs("ProjectName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
            'Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("id"), .Rows - 1, .ColIndex("id"))
        End With

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
       FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub



Private Sub ChangeLang()
    Me.Caption = "Search FollowAdminMeasure"

  '  XPLbl(1).Caption = "Employee Code"
  '  XPLbl(0).Caption = "Ekama No"
  '  XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
    XPLbl(2).Caption = "Location"
  '  XPLbl(6).Caption = "Status"
    XPLbl(9).Caption = "Name"
  '  XPLbl(3).Caption = "Nationality"
  '  XPLbl(7).Caption = "sponsor"
  '  XPLbl(9).Caption = "Name"
    XPLbl(4).Caption = "Job"
  '  XPLbl(5).Caption = "Religon"
'    Cmd1.Caption = "Print"
  '  ALLButton1.Caption = "Color Map"
Fra(2).Caption = "Process"
Fra(1).Caption = "Registration Date"
lbl(5).Caption = "From"
lbl(4).Caption = "From"
lbl(6).Caption = "To"
lbl(3).Caption = "To"
    With Me.fg
'        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
'        .TextMatrix(0, .ColIndex("ProfCode")) = "Employee Code"
'        .TextMatrix(0, .ColIndex("ProfNme")) = "Employee Name"
       ' .TextMatrix(0, .ColIndex("ProfPhone")) = "Employee Phone"

       ' .TextMatrix(0, .ColIndex("ProfID")) = "Employee N0"
'        .TextMatrix(0, .ColIndex("JobTypeName")) = " JobTypeName "
'        .TextMatrix(0, .ColIndex("DepartmentName")) = " Location "
       ' .TextMatrix(0, .ColIndex("nationality")) = " Nationality"
'        .TextMatrix(0, .ColIndex("datee")) = "  Date"
'        .TextMatrix(0, .ColIndex("id")) = " Process No "
'        .TextMatrix(0, .ColIndex("procedure")) = " Procedures "
       ' .TextMatrix(0, .ColIndex("kafelname")) = " Sponsor Name  "
     
    End With

End Sub


