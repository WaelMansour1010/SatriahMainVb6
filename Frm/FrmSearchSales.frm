VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSearchSales 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ „‰«œÌ» «·„»Ì⁄« "
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "FrmSearchSales.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   8100
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
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2700
      Width           =   1995
      Begin VB.TextBox txtcode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„ÊŸ›"
         Height          =   195
         Index           =   4
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»"
      Height          =   645
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   7995
      Begin VB.TextBox txtname 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   1515
      End
      Begin VB.TextBox txtbranch 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtgroup 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   6735
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   195
         Index           =   8
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Ã„Ê⁄Â"
         Height          =   195
         Index           =   0
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   3435
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      _cx             =   14102
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
      FormatString    =   $"FrmSearchSales.frx":038A
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
      TabIndex        =   6
      Top             =   4080
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   7
      Top             =   4080
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2940
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearchSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            GetData
        Case 1
            clear_all Me
        Case 2
            Unload Me
    End Select
End Sub
Private Sub Fg_EnterCell()

    On Error GoTo ErrTrap
  
    FrmPay_Garanty_Shipment.FindRec7 val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("empid")))
ErrTrap:
End Sub
Private Sub Form_Load()

    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set DCboSearch = New clsDCboSearch
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set GrdBack = New ClsBackGroundPic
    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    StrSQL = "SELECT dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode,"
    StrSQL = StrSQL & " dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    StrSQL = StrSQL & " dbo.TBLSalesRepGroups.id AS Expr1, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, dbo.TblEmpJobsTypes.JobTypeID,"
    StrSQL = StrSQL & " dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee"
    StrSQL = StrSQL & " FROM dbo.TBLSalesRepData INNER JOIN"
    StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData.JobID = dbo.TblEmpJobsTypes.JobTypeID"
    
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TBLSalesRepData.id >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLSalesRepData.id >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLSalesRepData.id <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLSalesRepData.id <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If

     If Me.txtcode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Code =" & Me.txtcode.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Code =" & Me.txtcode.Text & ""
        End If
    End If

    If Me.txtname.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.txtname.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.txtname.Text & "%'"
        End If
    End If
    
    If Me.txtbranch.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBranchesData.branch_name like '%" & Me.txtbranch.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBranchesData.branch_name like '%" & Me.txtbranch.Text & "%'"
        End If
    End If
    
    If Me.txtgroup.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLSalesRepGroups.name like '%" & Me.txtgroup.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLSalesRepGroups.name like '%" & Me.txtgroup.Text & "%'"
        End If
    End If

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TBLSalesRepData.id "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(i, .ColIndex("empid")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                    .TextMatrix(i, .ColIndex("branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                    .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                    .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                    .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    
    Me.Caption = "Search Sales Representative"

    Me.Fra(0).Caption = "By"
    lbl(4).Caption = "Emp Code"
    lbl(3).Caption = "Emp Name"
    lbl(5).Caption = "From"
    lbl(6).Caption = "To"
    lbl(0).Caption = "Group"
    lbl(8).Caption = "Branch"
    lbl(2).Caption = "Total"
    Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Process"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("ClientName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("group")) = "Group"
       .TextMatrix(0, .ColIndex("branch")) = "Branch"
    End With
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
End Sub
Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
End Sub
