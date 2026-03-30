VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSearchEmpSalary3 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ  Œ’Ì’ «·⁄„«·Â"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   Icon            =   "FrmSearchEmpSalary3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
      Height          =   645
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   5355
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   3
         Left            =   3855
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1335
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   7785
      Begin MSDataListLib.DataCombo dcopr 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo Dcterm 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„·Ì…"
         Height          =   315
         Index           =   4
         Left            =   6765
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·„‘—Ê⁄"
         Height          =   315
         Index           =   5
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»‰œ"
         Height          =   315
         Index           =   2
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   720
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _cx             =   18124
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchEmpSalary3.frx":038A
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
      TabIndex        =   1
      Top             =   4200
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
      TabIndex        =   2
      Top             =   4200
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
      TabIndex        =   3
      Top             =   4200
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
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   2295
   End
End
Attribute VB_Name = "FrmSearchEmpSalary3"
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
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub














Private Sub dcproject_Change()
dcproject_Click (0)

End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF3 Then
               FrmProjectSearch.lblSearchtype.Caption = 32
               FrmProjectSearch.show vbModal
        End If
End Sub

Private Sub Dcterm_Click(Area As Integer)


    Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
  If dcproject.BoundText <> "" Then
        
         If Me.Dcterm.BoundText <> "" Then
       '  Dcombos.GetProcessOfProjedt
         Dcombos.GetProcessOfProjedt dcopr, val(dcproject.BoundText), , val(Dcterm.BoundText), 2
         End If
       
    End If
End Sub
Private Sub dcproject_Click(Area As Integer)


    If dcproject.BoundText <> "" Then

        fillterms (val(dcproject.BoundText))
    End If

End Sub
Private Sub Dcterm_Change()
Dcterm_Click (0)
End Sub
Function fillterms(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.Dcterm, My_SQL
       
        
    dcopr.ReFill
End Function

Private Sub Fg_Click()
FrmEmpSalary3.Retrive (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
End Sub


Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  Set Dcombos = New ClsDataCombos
 

  
      Dim My_SQL As String
  If SystemOptions.UserInterface = ArabicInterface Then
      My_SQL = " select id,Project_name from projects where not(Project_name is null) and Project_name<>N'""' order by Project_name"
  Else
       My_SQL = " select id,Project_nameE from projects where not (Project_nameE is null) and Project_nameE<>N'""' order by Project_nameE"
  End If
    fill_combo dcproject, My_SQL
    My_SQL = "    select oprid,des from dbo.projects_des"

    fill_combo Me.Dcterm, My_SQL
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    'SetDtpickerDate Me.DtpDateFrom
    'SetDtpickerDate Me.DtpDateTo

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

StrSQL = "SELECT     dbo.opr_Employee.id, dbo.opr_Employee.Project_id, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.opr_Employee.PandID, dbo.projects_des.des, "
StrSQL = StrSQL & "                     dbo.opr_Employee.Years, dbo.opr_Employee.RecordDate, dbo.opr_Employee.Auto, dbo.opr_Employee.opr_type, dbo.opr_Employee.FromDate,"
StrSQL = StrSQL & "                      dbo.opr_Employee.ToDate , dbo.opr_Employee.OpraID, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
StrSQL = StrSQL & " FROM         dbo.opr_Employee LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProcessDEF ON dbo.opr_Employee.OpraID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.projects_des ON dbo.opr_Employee.PandID = dbo.projects_des.oprid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.projects ON dbo.opr_Employee.Project_id = dbo.projects.id"
    BolBegine = False
    StrWhere = ""

    '///////////////////
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.opr_Employee.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.opr_Employee.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.opr_Employee.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.opr_Employee.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
 
    If Me.dcproject.Text <> "" And (val(dcproject.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.opr_Employee.Project_id =" & Me.dcproject.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.opr_Employee.Project_id =" & Me.dcproject.BoundText & ""
        End If
    End If

       If Me.Dcterm.Text <> "" And (val(Dcterm.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.opr_Employee.PandID =" & Me.Dcterm.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.opr_Employee.PandID =" & Me.Dcterm.BoundText & ""
        End If
    End If
       If Me.dcopr.Text <> "" And (val(dcopr.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.opr_Employee.OpraID =" & Me.dcopr.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.opr_Employee.OpraID =" & Me.dcopr.BoundText & ""
        End If
    End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.opr_Employee.id "
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
                        
               ' If Not (IsNull(rs("RecordDate").value)) Then
               '     .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
               ' End If
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
            .TextMatrix(i, .ColIndex("ProcessName")) = IIf(IsNull(rs("ProcessName").value), "", rs("ProcessName").value)
          
            Else
            .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(rs("Project_nameE").value), "", rs("Project_nameE").value)
            .TextMatrix(i, .ColIndex("ProcessName")) = IIf(IsNull(rs("ProcessNameE").value), "", rs("ProcessNameE").value)
            
            End If
           .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
            
                
                
        
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Saerch Projects Labors Allocate"
'lbprocess.Caption = "No Transection"
lbl(5).Caption = " Project Name"
lbl(2).Caption = "Pand Name"
lbl(4).Caption = "Process Name"
lbprocess.Caption = "Saerch By"
lbl(3).Caption = "From"
'lbreg.Caption = "Date Transection"
lblLW.Caption = "Saerch By"
lbl(6).Caption = "To"

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
       
        .TextMatrix(0, .ColIndex("id")) = "TransID"
         .TextMatrix(0, .ColIndex("Project_name")) = " Project Name"
        .TextMatrix(0, .ColIndex("des")) = "Pand Name"
       .TextMatrix(0, .ColIndex("ProcessName")) = "Process Name"
    End With
  '
End Sub


