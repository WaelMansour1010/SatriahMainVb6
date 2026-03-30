VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSearchEqupment 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ «·„⁄œ«  / «·√·« "
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "FrmSearchEqupment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1335
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   6345
      Begin VB.TextBox TxtSearchEmp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "⁄›Ê« Ì—ÃÌ «œŒ«· ﬂÊœ «·„⁄œÂ ·Ê «·«·Â"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DcFixedAssets 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Tag             =   "Õœœ «”„ «·„⁄œ…"
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEmp1 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Tag             =   "Õœœ «”„ «·„⁄œ…"
         Top             =   600
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEmp2 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Tag             =   "Õœœ «”„ «·„⁄œ…"
         Top             =   960
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„”«⁄œ"
         Height          =   285
         Index           =   12
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„⁄œÂ"
         Height          =   285
         Index           =   3
         Left            =   5070
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„‘€·"
         Height          =   285
         Index           =   0
         Left            =   5070
         TabIndex        =   8
         Top             =   615
         Width           =   1125
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      _cx             =   15584
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchEqupment.frx":038A
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
      Top             =   3480
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
      Top             =   3480
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
      Top             =   3480
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
Attribute VB_Name = "FrmSearchEqupment"
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






Private Sub DCEmp1_Change()
DCEmp1_Click (0)
End Sub

Private Sub DCEmp1_Click(Area As Integer)
       If val(DCEmp1.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DCEmp1.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub

Private Sub DCEmp2_Click(Area As Integer)
    If val(DCEmp2.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DCEmp2.BoundText, EmpCode
    TxtSearchEmp.text = EmpCode
End Sub

Private Sub DcFixedAssets_Click(Area As Integer)
TxtCode.text = getFixedAsstName(val(DcFixedAssets.BoundText), "Code")
End Sub

Private Sub Fg_Click()
frmequipment.FindRec (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
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
  Dcombos.GetFixedAssets Me.DcFixedAssets, True
     Dcombos.GetEmployees Me.DCEmp1
     Dcombos.GetEmployees Me.DCEmp2
    
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

StrSQL = "SELECT     dbo.TblEquipments.name, dbo.TblEquipments.Code, dbo.TblEquipments.id, dbo.TblEquipments.Notes, dbo.TblEquipments.HourCount, "
StrSQL = StrSQL & "                      dbo.TblEquipments.fixedAssetid, dbo.FixedAssets.Name AS FixedName, dbo.FixedAssets.namee, dbo.TblEquipments.Stopvalue, dbo.TblEquipments.StopPercentage,"
StrSQL = StrSQL & "                       dbo.TblEquipments.empID1, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblEquipments.empID2,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name AS HelpEmp_Name, TblEmployee_1.Fullcode AS HelpFullcode, TblEmployee_1.Emp_Namee AS HelpEmp_NameE,"
StrSQL = StrSQL & "                       dbo.TblEquipments.helper, dbo.TblEquipments.Hourdipp, dbo.TblEquipments.Employer, dbo.TblEquipments.Rent, dbo.TblEquipments.ChKLockeq,"
StrSQL = StrSQL & "                       dbo.TblEquipments.[interval] , dbo.TblEquipments.UsedElectricPriceH, dbo.TblEquipments.UsedPowerPriceH"
StrSQL = StrSQL & "  FROM         dbo.TblEquipments LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblEquipments.empID2 = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblEquipments.empID1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.FixedAssets ON dbo.TblEquipments.fixedAssetid = dbo.FixedAssets.id"
    BolBegine = False
    StrWhere = ""

    '///////////////////
     If Me.TxtSearchCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND TblEmployee_2.fullcode ='" & Me.TxtSearchCode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where TblEmployee_2.fullcode ='" & Me.TxtSearchCode.text & "'"
        End If
    End If
        If Me.TxtCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEquipments.Code ='" & Me.TxtCode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEquipments.Code ='" & Me.TxtCode.text & "'"
        End If
    End If
         If Me.TxtSearchEmp.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND TblEmployee_1.Fullcode ='" & Me.TxtSearchEmp.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where TblEmployee_1.Fullcode ='" & Me.TxtSearchEmp.text & "'"
        End If
    End If
    
    If Me.DCEmp1.text <> "" And (val(DCEmp1.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEquipments.empID1 =" & Me.DCEmp1.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEquipments.empID1 =" & Me.DCEmp1.BoundText & ""
        End If
    End If
       If Me.DCEmp2.text <> "" And (val(DCEmp2.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEquipments.empID2 =" & Me.DCEmp2.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEquipments.empID2 =" & Me.DCEmp2.BoundText & ""
        End If
    End If
    


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblEquipments.id "
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
            .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("FixedName").value), "", rs("FixedName").value)
          .TextMatrix(i, .ColIndex("HelpEmp_Name")) = IIf(IsNull(rs("HelpEmp_Name").value), "", rs("HelpEmp_Name").value)
            Else
            .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            .TextMatrix(i, .ColIndex("HelpEmp_Name")) = IIf(IsNull(rs("HelpEmp_NameE").value), "", rs("HelpEmp_NameE").value)
            End If
           .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs("Code").value), "", rs("Code").value)
            
                
                
        
            
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
  Me.Caption = "Saerch Equipment / Machinery"
'lbprocess.Caption = "No Transection"
'lbl(5).Caption = "From"
lbl(12).Caption = "Assistant Name"
lbl(0).Caption = "Operator Name"
lbl(3).Caption = "Equipment"
'lbreg.Caption = "Date Transection"
lblLW.Caption = "Saerch By"
'lbl(2).Caption = "Total"

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("Code")) = "Equipment Code"
         .TextMatrix(0, .ColIndex("name")) = "Equipment Name"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Operator Name"
       .TextMatrix(0, .ColIndex("HelpEmp_Name")) = "Assistant Name"
    End With
  '
End Sub



Private Sub TxtCode_KeyPress(KeyAscii As Integer)
 Dim FixedID As Integer
If KeyAscii = vbKeyReturn Then
GetFixedIDFromCode TxtCode.text, FixedID
DcFixedAssets.BoundText = FixedID
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DCEmp1.BoundText = EmpID
    End If
    
End Sub

Private Sub TxtSearchEmp_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchEmp.text, EmpID
        DCEmp2.BoundText = EmpID
    End If
End Sub
