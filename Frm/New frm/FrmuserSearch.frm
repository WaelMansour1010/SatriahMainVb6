VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmUserSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·„” Œœ„Ì‰"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmuserSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   10080
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtCompanyName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2970
      Width           =   2955
   End
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
      Height          =   375
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3390
      Width           =   2385
   End
   Begin VB.TextBox XPTxtComID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2610
      Width           =   2955
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2505
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      _cx             =   17754
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmuserSearch.frx":030A
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
      Left            =   2010
      TabIndex        =   5
      Top             =   3510
      Width           =   915
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
      Height          =   375
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Top             =   3510
      Width           =   915
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
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   7
      Top             =   3510
      Width           =   915
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
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   345
      Index           =   2
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label lblSearchtype 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   345
      Index           =   0
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2955
      Width           =   1185
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„” Œœ„"
      Height          =   315
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2610
      Width           =   1185
   End
End
Attribute VB_Name = "FrmUserSearch"
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
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.Rows = 1

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
    
        If Me.lblSearchtype.Caption = 0 Then
            FrmEditUsers.FindRec val(Fg.TextMatrix(Fg.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 1 Then
            FrmPermission.TxtCode = (Fg.TextMatrix(Fg.Row, 3))
  
   



        End If

    End If
Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Cmd_Click (2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not Fg.TextMatrix(Fg.Row, Fg.ColIndex("Code")) = "" Then
            Fg_Click
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
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos

Dcombos.GetBranches Me.dcBranch

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    Exit Sub
ErrTrap:

End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("UserID")) = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
                .TextMatrix(Num, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", (rs("UserName").value))
                .TextMatrix(Num, .ColIndex("Emp_Code")) = IIf(IsNull(rs("Emp_Code").value), "", (rs("Emp_Code").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                     .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
                Else
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", Trim(rs("branch_namee").value))
                End If
             End With

            rs.MoveNext
        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
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
    StrSQL = "SELECT     dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.TblUsers.Empid, dbo.TblEmployee.Emp_Code, dbo.TblUsers.BranchId, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "   dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & "  FROM         dbo.TblUsers LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.TblUsers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"

StrSQL = StrSQL & "  where 1=1  "
 
    Begin = True

    If XPTxtComID.text <> "" Then
         
            StrWhere = StrWhere + " AND (dbo.TblEmployee.Emp_Code LIKE '%" & XPTxtComID & "%')  "
      
    End If

    If TxtCompanyName.text <> "" Then
            
                StrWhere = StrWhere + "  AND (dbo.TblUsers.UserName LIKE '%" & TxtCompanyName & "%')"
    End If
If dcBranch.text <> "" Then
       StrWhere = StrWhere + " and   (dbo.TblUsers.BranchId = " & val(dcBranch.BoundText) & ") "
End If

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub ChangeLang()
    Me.Caption = "Supplier Search..."
    XPLbl(1).Caption = "Supplier Code"
    XPLbl(0).Caption = "Supplier Name"
    XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Count")) = "Serial"
        .TextMatrix(0, .ColIndex("Code")) = "Supplier Code"
        .TextMatrix(0, .ColIndex("Name")) = "Supplier Name"
        .TextMatrix(0, .ColIndex("Phone")) = "Supplier Phone"
        .TextMatrix(0, .ColIndex("Mobile")) = "Supplier Mobile"
    
    End With

End Sub

