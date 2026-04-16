VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEmpMoveSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ Œÿ«»  ⁄—Ì›"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   Icon            =   "FrmEmpMoveSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   11955
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3840
      Width           =   9075
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄—› „«·Ì"
         Height          =   195
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄—Ì›"
         Height          =   195
         Index           =   1
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   345
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   62259203
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   345
         Left            =   6240
         TabIndex        =   31
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   62259203
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   9
         Left            =   5460
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   7
         Left            =   8295
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2580
      Width           =   5115
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   915
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Index           =   11
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Õ”»"
      Height          =   645
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   11715
      Begin VB.TextBox TxtFileNo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox TxtIqama 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1875
      End
      Begin VB.TextBox TxtJob 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5700
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox TxtNationality 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„·›"
         Height          =   195
         Index           =   4
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   300
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«ﬁ«„Â"
         Height          =   195
         Index           =   3
         Left            =   10575
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   195
         Index           =   8
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   195
         Index           =   0
         Left            =   4935
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·ÌÂ"
      Height          =   645
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
      Width           =   3795
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
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   3960
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
      TabIndex        =   6
      Top             =   3960
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
      TabIndex        =   7
      Top             =   3960
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   21
      Top             =   0
      Width           =   11835
      _cx             =   20876
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmEmpMoveSearch.frx":038A
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmEmpMoveSearch"
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


Private Sub Fg_Click()

   
 FrmDefineEmp.Retrive (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))


End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmpName
    ' Dcombos.GetClientName Me.DCEmp_Name
    Set DCboSearch = New clsDCboSearch
    'Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
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

    StrSQL = " SELECT     dbo.TblDefin.RecordDate, dbo.TblDefin.RecordDateH, dbo.TblDefin.ID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDefin.EmpID, "
     StrSQL = StrSQL & "                 dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
  StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3,"
   StrSQL = StrSQL & "                    dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.TblDefin.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
  StrSQL = StrSQL & "                     dbo.TblDefin.TypeDef, dbo.TblDefin.FileNo, dbo.TblDefin.CurrSala, dbo.TblDefin.PointSala, dbo.TblDefin.Iqama, dbo.TblDefin.IssuDate, dbo.TblDefin.IssuDateH,"
 StrSQL = StrSQL & "                      dbo.TblDefin.Remark , dbo.TblDefin.UserID, dbo.TblEmployee.Nationality, dbo.TblEmpJobsTypes.JobTypeID"
StrSQL = StrSQL & "  FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblDefin ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblDefin.BranchID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblDefin.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblDefin.BranchID = dbo.TblBranchesData.branch_id"

    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblDefin.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
 

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.ID<=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
     If Me.TxtSearchCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode '%" & Me.TxtSearchCode.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.text & "%'"
        End If
    End If
'////////////////////////
 If Me.TxtIqama.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.Iqama like '%" & Me.TxtIqama.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.Iqama like '%" & Me.TxtIqama.text & "%'"
        End If
    End If
    ''/////////////
     If Me.TxtJob.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpJobsTypes.JobTypeName like '%" & Me.TxtJob.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpJobsTypes.JobTypeName like '%" & Me.TxtJob.text & "%'"
        End If
    End If
    '''./////////
     If Me.TxtNationality.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEmployee.Nationality like '%" & Me.TxtNationality.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEmployee.Nationality like '%" & Me.TxtNationality.text & "%'"
        End If
    End If
    '''///////////////
     If Me.TxtFileNo.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.FileNolike '%" & Me.TxtFileNo.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.FileNo like '%" & Me.TxtFileNo.text & "%'"
        End If
    End If
    ''///////////
   
   If Me.DcboEmpName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.EmpID =" & Me.DcboEmpName.BoundText & ""
        End If
    End If
      If Me.Opt(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.TypeDef = 0 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.TypeDef = 0 "
        End If
    End If
If Me.Opt(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.TypeDef = 1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.TypeDef = 1"
        End If
    End If
 

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDefin.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDefin.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblDefin.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblDefin.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblDefin.ID "
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
         
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
             .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("iqqmq")) = IIf(IsNull(rs("Iqama").value), "", rs("Iqama").value)
                .TextMatrix(i, .ColIndex("job")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                .TextMatrix(i, .ColIndex("nationality")) = IIf(IsNull(rs("Nationality").value), "", rs("Nationality").value)
               .TextMatrix(i, .ColIndex("fileno")) = IIf(IsNull(rs("FileNo").value), "", rs("FileNo").value)
              If rs("TypeDef").value = False Then
                .TextMatrix(i, .ColIndex("TypeDef")) = " ⁄—Ì› „«·Ì"
                Else
             .TextMatrix(i, .ColIndex("TypeDef")) = " ⁄—Ì›"
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
  Me.Caption = "Search Definition Employee's Speech"
Me.lbprocess.Caption = "Process No"
lbl(7).Caption = "From"
lbl(9).Caption = "To"
lbl(11).Caption = "Emp"
lbl(2).Caption = "Total"
lbl(3).Caption = "IqamaNo"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Nationality"
lbl(8).Caption = "JobName"

lbl(4).Caption = "FileNo"
lbl(7).Caption = "Storekeeper"
'Me.lbreg.Caption = "Date Registration"
Me.Opt(0).Caption = "Define Financial"
Me.Opt(1).Caption = "Definition"
Me.Opt(0).RightToLeft = False
Me.Opt(1).RightToLeft = False

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Proce"
        .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
         .TextMatrix(0, .ColIndex("name")) = "EmpName"
        .TextMatrix(0, .ColIndex("iqqmq")) = "Iqama No"
       .TextMatrix(0, .ColIndex("job")) = "JobName"
       .TextMatrix(0, .ColIndex("nationality")) = "Nationality "
         .TextMatrix(0, .ColIndex("fileno")) = "FileNo"
        .TextMatrix(0, .ColIndex("TypeDef")) = "TypeDef"
       '.TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub


Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
