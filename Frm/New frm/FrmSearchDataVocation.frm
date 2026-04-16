VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchDataVocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ »Ì«‰«  «·«Ã«“…"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "FrmSearchDataVocation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   10980
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
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   11025
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«÷—«—Ì…"
         Height          =   255
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—”„Ì…"
         Height          =   255
         Index           =   0
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtNumEkama 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   600
         Width           =   4275
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   5760
         TabIndex        =   21
         Top             =   600
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcemplocation 
         Height          =   315
         Left            =   5760
         TabIndex        =   23
         Top             =   960
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboJobsType 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·«Ã«“Â"
         Height          =   285
         Index           =   12
         Left            =   4200
         TabIndex        =   31
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«ﬁ«„Â"
         Height          =   285
         Index           =   11
         Left            =   4200
         TabIndex        =   28
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›Â"
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   26
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Êﬁ⁄ «·⁄„·"
         Height          =   285
         Index           =   8
         Left            =   9600
         TabIndex        =   24
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—Â"
         Height          =   285
         Index           =   7
         Left            =   9630
         TabIndex        =   22
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   0
         Left            =   9750
         TabIndex        =   20
         Top             =   255
         Width           =   1125
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·⁄„·ÌÂ"
      Height          =   555
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   4575
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2040
         TabIndex        =   6
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94240771
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94240771
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·ÌÂ"
      Height          =   645
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
      Width           =   3555
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   360
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
         Left            =   2775
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
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10995
      _cx             =   19394
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchDataVocation.frx":038A
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
      TabIndex        =   11
      Top             =   4800
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
      TabIndex        =   12
      Top             =   4800
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
      TabIndex        =   13
      Top             =   4800
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3060
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearchDataVocation"
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



Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    End Sub

Private Sub Fg_Click()
'FrmHolidayData.Retrive (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
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
    
    Dcombos.GetEmpLocations Me.dcemplocation
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpDepartments Me.DcbDept
    
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

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

StrSQL = "SELECT     dbo.tblHolidayData.id, dbo.tblHolidayData.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblHolidayData.Emp_ID, "
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.tblHolidayData.recorddate, dbo.tblHolidayData.DeparmentID,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.tblHolidayData.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "                     dbo.TblEmpJobsTypes.JobTypeNamee, dbo.tblHolidayData.visano, dbo.tblHolidayData.noofmonth, dbo.tblHolidayData.Returnbeforedate,"
StrSQL = StrSQL & "                       dbo.tblHolidayData.ReturnbeforedateH, dbo.tblHolidayData.DeparDate, dbo.tblHolidayData.DeparDateH, dbo.tblHolidayData.ExpectedReturndate,"
StrSQL = StrSQL & "                       dbo.tblHolidayData.ExpectedReturndateH, dbo.tblHolidayData.Remark, dbo.tblHolidayData.PostedDate, dbo.tblHolidayData.chekdate,"
StrSQL = StrSQL & "                       dbo.tblHolidayData.SpecificHolidyaType1 , dbo.tblHolidayData.ProjectID, dbo.EmpGroupDep.GroupName , dbo.tblHolidayData.NumEkama"
StrSQL = StrSQL & "  FROM         dbo.tblHolidayData LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.EmpGroupDep ON dbo.tblHolidayData.ProjectID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpJobsTypes ON dbo.tblHolidayData.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpDepartments ON dbo.tblHolidayData.DeparmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.tblHolidayData.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.tblHolidayData.branch_no = dbo.TblBranchesData.branch_id"

    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.tblHolidayData.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
      If RegCK(0).value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.tblHolidayData.SpecificHolidyaType1 =0 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.SpecificHolidyaType1 = 0 "
        End If
    End If
          If RegCK(1).value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.tblHolidayData.SpecificHolidyaType1 =1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.SpecificHolidyaType1 = 1 "
        End If
    End If

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    '///////////////////
         If Me.TxtNumEkama.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.NumEkama ='" & Me.TxtNumEkama.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.NumEkama ='" & Me.TxtNumEkama.Text & "'"
        End If
    End If
    
     If Me.TxtSearchCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.fullcode ='" & Me.TxtSearchCode.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.fullcode ='" & Me.TxtSearchCode.Text & "'"
        End If
    End If
    If Me.DcboEmpName.Text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.Emp_ID =" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.Emp_ID =" & Me.DcboEmpName.BoundText & ""
        End If
    End If
    
        If Me.DcbDept.Text <> "" And (val(DcbDept.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.DeparmentID =" & Me.DcbDept.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.DeparmentID =" & Me.DcbDept.BoundText & ""
        End If
    End If
         If Me.dcemplocation.Text <> "" And (val(dcemplocation.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.ProjectID =" & Me.dcemplocation.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.ProjectID =" & Me.dcemplocation.BoundText & ""
        End If
    End If
    If Me.DcboJobsType.Text <> "" And (val(DcboJobsType.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.JobTypeID =" & Me.DcboJobsType.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.JobTypeID =" & Me.DcboJobsType.BoundText & ""
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblHolidayData.recorddate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblHolidayData.recorddate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.tblHolidayData.recorddate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.tblHolidayData.recorddate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.tblHolidayData.ID"
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
               
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("recorddate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("recorddate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
            .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
          
            Else
            .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
            .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
            End If
           .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
           .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
           If SystemOptions.UserInterface = ArabicInterface Then
               If rs("SpecificHolidyaType1").value = True Then
               .TextMatrix(i, .ColIndex("typevocation")) = "«÷ÿ—«—ÌÂ"
               Else
                 .TextMatrix(i, .ColIndex("typevocation")) = "—”„ÌÂ"
               End If
               Else
                If rs("SpecificHolidyaType1").value = True Then
               .TextMatrix(i, .ColIndex("typevocation")) = "Forced"
               Else
                 .TextMatrix(i, .ColIndex("typevocation")) = "Official"
               End If
               End If

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
  Me.Caption = "Saerch Data of Vacation"
lbprocess.Caption = "No Transection"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbreg.Caption = "Date Transection"
lblLW.Caption = "Saerch By"
lbl(2).Caption = "Total"
lbl(0).Caption = "Employee"
lbl(7).Caption = "Department"
lbl(8).Caption = "Location"
lbl(9).Caption = "Job"
lbl(11).Caption = "Num Iqama"
lbl(12).Caption = "Type Vocation"
RegCK(0).RightToLeft = False
RegCK(1).RightToLeft = False
RegCK(0).Caption = "Official"
RegCK(1).Caption = "Forced"


     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("DepartmentName")) = "DepartmentName"
        .TextMatrix(0, .ColIndex("empname")) = "Emp Name"
       .TextMatrix(0, .ColIndex("GroupName")) = "Location"
         .TextMatrix(0, .ColIndex("JobTypeName")) = "Job Name"
       .TextMatrix(0, .ColIndex("NumEkama")) = "Num Iqama"
       .TextMatrix(0, .ColIndex("typevocation")) = "Type Vocation"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
End Sub
