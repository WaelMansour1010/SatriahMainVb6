VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCatchSearsh 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ ”‰œ ﬁ»÷ ⁄«„"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmCatchSearsh.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   795
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·”‰œ"
      Height          =   795
      Index           =   0
      Left            =   -240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   4335
      Begin MSComCtl2.DTPicker StarDateFrom 
         Height          =   330
         Left            =   2370
         TabIndex        =   11
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   268304387
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker StarDateTo 
         Height          =   330
         Left            =   450
         TabIndex        =   12
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   268304387
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   8
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   7
         Left            =   1815
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2760
      Width           =   4455
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _cx             =   17701
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCatchSearsh.frx":038A
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
      Left            =   8130
      TabIndex        =   1
      Top             =   4320
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
      Left            =   7290
      TabIndex        =   2
      Top             =   4320
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
      Left            =   6510
      TabIndex        =   3
      Top             =   4320
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   4200
      TabIndex        =   15
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcGeneralBox 
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   3480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo Dcbranch 
      Height          =   315
      Left            =   4200
      TabIndex        =   19
      Top             =   3120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   195
      Index           =   3
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3960
      Width           =   570
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   285
      Index           =   17
      Left            =   9105
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3150
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰œÊﬁ «·›—⁄Ì"
      Height          =   285
      Index           =   14
      Left            =   8730
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3870
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰œÊﬁ «·—∆Ì”Ì"
      Height          =   285
      Index           =   15
      Left            =   8445
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3480
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   195
      Index           =   9
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2820
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
      TabIndex        =   6
      Top             =   2820
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   195
      Index           =   0
      Left            =   11115
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3630
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   435
      Index           =   10
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   2055
   End
End
Attribute VB_Name = "FrmCatchSearsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
            GetData

        Case 1
            clear_all Me
Me.StarDateFrom.value = ""
Me.StarDateTo.value = ""

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub












Private Sub fg_Click()

    With Me.FG

        If .row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.row, .ColIndex("id"))) = 0 Then
            Exit Sub
        End If
 FrmBankDeposite3.Retrive val(.TextMatrix(.row, .ColIndex("id")))
       
    

    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub ChangeLang()
    'Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
'    Label1.Visible = False

    'Cmd(0).Caption = "New"
    'Cmd(1).Caption = "Edit"
    'Cmd(2).Caption = "Save"
    'Cmd(3).Caption = "Undo"
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
 'Cmd(9).Caption = "Print"
    Cmd(2).Caption = "Exit"
 '   CmdHelp.Caption = "Help"

    Me.Caption = " Search For Distribution Expenses on Items "
    'EleHeader.Caption = Me.Caption
    lbl(5).Caption = "From"
    lbl(6).Caption = "To"
    lbl(4).Caption = "From"
    lbl(3).Caption = "To "
   ' Frame10.Caption = "Select Store"
    Fra(2).Caption = "Registration No  "
    Fra(1).Caption = "Registration Date  "
lbl(7).Caption = "AccountName"
lbl(8).Caption = "GroupName"
lbl(0).Caption = "ItemName"
lbl(2).Caption = "Total"
   'lbl(8).Caption = "By"
   ' lbl(7).Caption = "Curr rec."
   ' lbl(6).Caption = "rec. count"

   With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("ID")) = "NO"
        .TextMatrix(0, .ColIndex("Account_Name")) = "AccountName"
         .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
        .TextMatrix(0, .ColIndex("GroupName")) = "GroupName"
         .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("TypeValue")) = "TypeValue"
        .TextMatrix(0, .ColIndex("Vlue")) = "Value"
        .TextMatrix(0, .ColIndex("RemarkD3")) = "Remarks"

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
  
    Set DCboSearch = New clsDCboSearch
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBoxes Me.DcGeneralBox
    Dcombos.GetBranches Me.dcBranch
  
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

StarDateFrom.value = Date
StarDateTo.value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = "SELECT     TOP 100 PERCENT dbo.tblGeneralCashing.id, dbo.tblGeneralCashing.RecordDate, dbo.tblGeneralCashing.NoteSerial1, dbo.tblGeneralCashing.NoteSerial, "
StrSQL = StrSQL & "                      dbo.tblGeneralCashing.OldNoteSerial1, dbo.tblGeneralCashing.Remarks, dbo.tblGeneralCashing.ToDate, dbo.tblGeneralCashing.FromDate,"
StrSQL = StrSQL & "                       dbo.tblGeneralCashing.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblGeneralCashing.GeneralBoxId,"
 StrSQL = StrSQL & "                      dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE, dbo.tblGeneralCashing.SubBoxId, TblBoxesData_1.BoxName AS SubDcboBox,"
 StrSQL = StrSQL & "                      TblBoxesData_1.BoxNameE AS SubDcboBoxE, dbo.tblGeneralCashing.NoteID"
StrSQL = StrSQL & "  FROM         dbo.tblGeneralCashing LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBoxesData TblBoxesData_1 ON dbo.tblGeneralCashing.SubBoxId = TblBoxesData_1.BoxID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblBoxesData ON dbo.tblGeneralCashing.GeneralBoxId = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblGeneralCashing.branch_no = dbo.TblBranchesData.branch_id"

 
    BolBegine = False
    StrWhere = ""



    If val(Me.dcBranch.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblGeneralCashing.branch_no=" & Me.dcBranch.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblGeneralCashing.branch_no =" & Me.dcBranch.BoundText & ""
        End If
    End If
    

    If val(Me.DcGeneralBox.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblGeneralCashing.GeneralBoxId=" & Me.DcGeneralBox.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblGeneralCashing.GeneralBoxId =" & Me.DcGeneralBox.BoundText & ""
        End If
    End If
 If val(Me.DcboBox.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblGeneralCashing.SubBoxId=" & Me.DcboBox.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblGeneralCashing.SubBoxId =" & Me.DcboBox.BoundText & ""
        End If
    End If
    
    If Me.TxtNoteSerial.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblGeneralCashing.NoteSerial1 like '%" & Me.TxtNoteSerial.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.tblGeneralCashing.NoteSerial1 like '%" & Me.TxtNoteSerial.text & "%'"
        End If
    End If
   If Me.txtremark.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblGeneralCashing.Remarks like '%" & Me.txtremark.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.tblGeneralCashing.Remarks like '%" & Me.txtremark.text & "%'"
        End If
    End If


    If Not IsNull(Me.StarDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.tblGeneralCashing.RecordDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.tblGeneralCashing.RecordDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.StarDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.tblGeneralCashing.RecordDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.tblGeneralCashing.RecordDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        End If
    End If
    

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.tblGeneralCashing.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                 .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                 .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                                
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                 .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                 .TextMatrix(i, .ColIndex("SubDcboBox")) = IIf(IsNull(rs("SubDcboBox").value), "", rs("SubDcboBox").value)
               
                Else
                 .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                 .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                 .TextMatrix(i, .ColIndex("SubDcboBox")) = IIf(IsNull(rs("SubDcboBoxE").value), "", rs("SubDcboBoxE").value)
                End If
             
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub






