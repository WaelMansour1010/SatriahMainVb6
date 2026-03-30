VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmEnserviceSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·» ‰Â«Ì… «·Œœ„…"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "FrmENdServiceSearsh.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1395
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3240
      Width           =   2175
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98631683
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98631683
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   9
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   7
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   4275
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   3495
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1485
      Index           =   0
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   5115
      Begin VB.TextBox txtname 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   3555
      End
      Begin VB.TextBox txtcode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin MSDataListLib.DataCombo dctype 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ ‰Â«Ì… «·Œœ„…"
         Height          =   195
         Index           =   4
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ã“¡ „‰ «”„ «·„ÊŸ›"
         Height          =   195
         Index           =   8
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Index           =   0
         Left            =   3975
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1020
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _cx             =   12832
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
      FormatString    =   $"FrmENdServiceSearsh.frx":038A
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
      Left            =   2490
      TabIndex        =   1
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
      Left            =   1650
      TabIndex        =   2
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
      Left            =   870
      TabIndex        =   3
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
      Caption         =   "«·„ÊŸ›"
      Height          =   195
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   2820
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
      Width           =   2775
   End
End
Attribute VB_Name = "FrmEnserviceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Index As Integer
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

 GetData
           
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub fg_Click()
     'On Error GoTo ErrTrap
  '  FrmModels.FindRec
  If Index = 0 Then
   End_oF_service.Retrive val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
   ElseIf Index = 1 Then
   FrmTypeExchange.txtTransaction_ID = val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
      ElseIf Index = 2 Then
   FrmPayments.TxtEndService.Text = val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
   End If
'ErrTrap:
                
            
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos

 Dcombos.GetJobEndService dctype
    Dcombos.GetEmployees Me.DcboEmp
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
   DtpDateFrom.value = ""
   Me.DtpDateTo.value = ""

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

StrSQL = "SELECT     dbo.End_of_service.opr_date, dbo.End_of_service.sal, dbo.End_of_service.sakn, dbo.End_of_service.bus, dbo.End_of_service.another, dbo.End_of_service.food, "
StrSQL = StrSQL & "                      dbo.End_of_service.mob, dbo.End_of_service.mang, dbo.End_of_service.total_salary, dbo.End_of_service.start_date, dbo.End_of_service.[end _date],"
StrSQL = StrSQL & "                        dbo.End_of_service.daycount, dbo.End_of_service.monthcount, dbo.End_of_service.yearcount, dbo.End_of_service.total, dbo.End_of_service.opr,"
StrSQL = StrSQL & "                        dbo.End_of_service.num, dbo.End_of_service.net, dbo.End_of_service.sal1, dbo.End_of_service.sakn1, dbo.End_of_service.bus1, dbo.End_of_service.another1,"
StrSQL = StrSQL & "                        dbo.End_of_service.food1, dbo.End_of_service.mob1, dbo.End_of_service.mang1, dbo.End_of_service.record_date, dbo.End_of_service.Type, dbo.jopstatus.name,"
 StrSQL = StrSQL & "                       dbo.jopstatus.namee, dbo.jopstatus.resignationInt, dbo.jopstatus.Vacation, dbo.End_of_service.Reaons, dbo.End_of_service.Des, dbo.End_of_service.TotalAdvance,"
 StrSQL = StrSQL & "                       dbo.End_of_service.TotalCash, dbo.End_of_service.LastTotal, dbo.End_of_service.VwithoutSa, dbo.End_of_service.TxtVlueVaction, dbo.TblUsers.UserName,"
StrSQL = StrSQL & "                        dbo.End_of_service.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
StrSQL = StrSQL & "                        dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.jopstatus.id,"
 StrSQL = StrSQL & "                       dbo.End_of_service.id AS idM"
StrSQL = StrSQL & "   FROM         dbo.jopstatus RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                       dbo.End_of_service INNER JOIN"
  StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.End_of_service.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                       dbo.TblUsers ON dbo.End_of_service.UserID = dbo.TblUsers.UserID ON dbo.jopstatus.id = dbo.End_of_service.Type"

    BolBegine = False
    StrWhere = "  "
If Index <> 0 Then
StrWhere = StrWhere & "where   not (NoteSerial is null) "
End If

 If CheckAprroveScreen("End_oF_service") = True And Index <> 0 Then
     StrWhere = StrWhere & " and approved =1"

End If

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "  AND dbo.End_of_service.id >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.End_of_service.id >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
 

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.End_of_service.id <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.End_of_service.id <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If


'////////////////////////
 If Me.txtName.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.txtName.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.txtName.Text & "%'"
        End If
    End If
   If Me.TxtCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode ='" & TxtCode.Text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Fullcode='" & TxtCode.Text & "'"
        End If
    End If
 If Me.DcboEmp.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblEmployee.Emp_ID =" & DcboEmp.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_ID=" & DcboEmp.BoundText & ""
        End If
    End If

 If Me.dctype.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.End_of_serviceDetails.IDMofrd =" & dctype.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.End_of_serviceDetails.IDMofrd=" & dctype.BoundText & ""
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.End_of_service.record_date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.End_of_service.record_date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.End_of_service.record_date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.End_of_service.record_date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.End_of_service.id"
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
             
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("idM").value), "", rs("idM").value)
      
            .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
            .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                  If Not (IsNull(rs("record_date").value)) Then
                    .TextMatrix(i, .ColIndex("recorddate")) = Format(rs("record_date").value, "yyyy/M/d")
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
  Me.Caption = "Search End_Service"
lbl(4).Caption = "Type"
'Me.LblClientName.Caption = "AsestName"
lbl(0).Caption = "Employee"
lbl(8).Caption = "Part Name"
lbl(2).Caption = "Total"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbreg.Caption = "Reg Date"
lbl(9).Caption = "From"
lbl(7).Caption = "To"
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("code")) = "EmpCode"
        .TextMatrix(0, .ColIndex("id")) = "No"
         .TextMatrix(0, .ColIndex("ClientName")) = "Employee"
        .TextMatrix(0, .ColIndex("type")) = "Type"
      .TextMatrix(0, .ColIndex("recorddate")) = "RecordeDate"
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

