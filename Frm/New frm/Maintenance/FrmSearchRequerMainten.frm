VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchRequerMainten 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·»«  «·’Ì«‰…"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   Icon            =   "FrmSearchRequerMainten.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   10695
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
   Begin VB.ComboBox ProblemTimID1 
      Height          =   315
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Index           =   1
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2700
      Width           =   4125
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   345
         Left            =   2100
         TabIndex        =   23
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   104267777
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   375
         Left            =   60
         TabIndex        =   24
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   104267777
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   315
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1485
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   10635
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox ProblemTimID 
         Height          =   315
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmSearchRequerMainten.frx":038A
         Height          =   315
         Left            =   8160
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
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
      Begin MSDataListLib.DataCombo DcbUnit 
         Bindings        =   "FrmSearchRequerMainten.frx":039F
         Height          =   315
         Left            =   5520
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
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
      Begin MSDataListLib.DataCombo DcbEquepment 
         Height          =   315
         Left            =   3720
         TabIndex        =   18
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Êﬁ  «·„‘ﬂ·…"
         Height          =   255
         Index           =   7
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„⁄œÂ"
         Height          =   255
         Index           =   4
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         Height          =   255
         Index           =   3
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   7200
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
      Width           =   10635
      _cx             =   18759
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
      FormatString    =   $"FrmSearchRequerMainten.frx":03B4
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
      Top             =   5160
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
      Top             =   5160
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
      Top             =   5160
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
   Begin VB.Label lbltype 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
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
Attribute VB_Name = "FrmSearchRequerMainten"
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
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub Fg_EnterCell()
  On Error GoTo ErrTrap
  If lbltype.Caption = 0 Then
    FrmRequerMainten.Retrive val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
  ElseIf lbltype.Caption = 1 Then
  FrmOrderMaintin.TxtOrder = val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
  End If
ErrTrap:

End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
  Dcombos.GetEmpDepartments Me.DcbUnit
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEquipments DcbEquepment
    
    '  If SystemOptions.UserInterface = EnglishInterface Then
     
    '  Me.DcbOrderStatus.AddItem "New"
    '    Me.DcbOrderStatus.AddItem "Accept Customer"
    '    Me.DcbOrderStatus.AddItem "Final Maintenance"

       If SystemOptions.UserInterface = EnglishInterface Then
            ProblemTimID.AddItem "During Production"
            ProblemTimID.AddItem "During Start up"
            ProblemTimID.AddItem "During Repair"
            ProblemTimID.AddItem "Others"
            ProblemTimID1.AddItem "During Production"
            ProblemTimID1.AddItem "During Start up"
            ProblemTimID1.AddItem "During Repair"
            ProblemTimID1.AddItem "Others"
     '   SetInterface Me
     '   ChangeLang
        Else
        ProblemTimID.AddItem "«À‰«¡ «· ’‰Ì⁄"
        ProblemTimID.AddItem "«À‰«¡ »œ¡ «· ‘€Ì·"
        ProblemTimID.AddItem "«À‰«¡ «·«’·«Õ"
        ProblemTimID.AddItem "«Œ—Ï"
            ProblemTimID1.AddItem "«À‰«¡ «· ’‰Ì⁄"
        ProblemTimID1.AddItem "«À‰«¡ »œ¡ «· ‘€Ì·"
        ProblemTimID1.AddItem "«À‰«¡ «·«’·«Õ"
        ProblemTimID1.AddItem "«Œ—Ï"

    End If
    Set DCboSearch = New clsDCboSearch
   ' Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
'    SetDtpickerDate Me.DtpDateFrom
'    SetDtpickerDate Me.DtpDateTo

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
    On Error Resume Next
StrSQL = " SELECT     dbo.TblRequerMainten.ID, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime, "
StrSQL = StrSQL & "                      dbo.TblRequerMainten.StartTime, dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate,"
StrSQL = StrSQL & "                        dbo.TblRequerMainten.StopDate, dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
StrSQL = StrSQL & "                        dbo.TblRequerMainten.EquepID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                        dbo.TblBranchesData.branch_nameE , dbo.FixedAssets.namee"
StrSQL = StrSQL & "   FROM         dbo.TblRequerMainten LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID"
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
   

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
   If Me.Dcbranch.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblRequerMainten.BranchID=" & Me.Dcbranch.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblRequerMainten.BranchID=" & Me.Dcbranch.BoundText & ""
        End If
    End If

   If Me.DcbUnit.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblRequerMainten.UnitID=" & Me.DcbUnit.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblRequerMainten.UnitID=" & Me.DcbUnit.BoundText & ""
        End If
    End If
  
    If Me.DcbEquepment.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblRequerMainten.EquepID=" & Me.DcbEquepment.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblRequerMainten.EquepID=" & Me.DcbEquepment.BoundText & ""
        End If
    End If
If val(Me.ProblemTimID.ListIndex) <> -1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblRequerMainten.ProblemTimID=" & Me.ProblemTimID.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblRequerMainten.ProblemTimID=" & Me.ProblemTimID.ListIndex & ""
        End If
    End If

   If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
           StrWhere = " Where  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblRequerMainten.id "
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
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                Else
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                End If
      

                ProblemTimID1.ListIndex = val(IIf(IsNull(rs("ProblemTimID").value), -1, rs("ProblemTimID").value))
                .TextMatrix(i, .ColIndex("ProblemTimID")) = ProblemTimID1.Text
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
    
  Me.Caption = "Search Requer Maintenance"
lbprocess.Caption = "OrderNo"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(11).Caption = "From"
lbl(0).Caption = "To"
Fra(1).Caption = "Date"
lblbr.Caption = "Branch"
lbl(4).Caption = "Machine"
lbl(3).Caption = "Dept"
lbl(7).Caption = "Time Problem"

lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date Registration"

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "OrderNo "
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("branch_name")) = "BranchName"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "DepartmentName"
       .TextMatrix(0, .ColIndex("Name")) = "Machine"
        .TextMatrix(0, .ColIndex("ProblemTimID")) = "TimeProblem"
       
    End With
  '
End Sub

Private Sub Text2_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = 2))) AND  dbo.TblCarsData.Fullcode like '%" & Text2.Text & "%'  "


   Dcombos.GetEquipments Me.DcbEquepment, str
   
    


End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

