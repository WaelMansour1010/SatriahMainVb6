VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAssestSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÃÀ ⁄‰ ‰ﬁ· Ê ”·Ì„ «·⁄Âœ"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   Icon            =   "FrmAsseetSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtfullcode 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”·Ì„"
      Height          =   1035
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2640
      Width           =   2295
      Begin MSComCtl2.DTPicker DtpDateFromDr 
         Height          =   330
         Left            =   90
         TabIndex        =   29
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64094211
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateToDr 
         Height          =   330
         Left            =   90
         TabIndex        =   30
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64094211
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   9
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   7
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.TextBox TxtFitter 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   2955
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   12435
      Begin VB.TextBox txtremarks 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   5355
      End
      Begin VB.TextBox txtqunatity 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7380
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtassest 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   195
         Index           =   12
         Left            =   5955
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„ÌÂ"
         Height          =   195
         Index           =   8
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄ÂœÂ"
         Height          =   195
         Index           =   0
         Left            =   11175
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·ÿ·»"
      Height          =   1035
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64094211
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   64094211
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   645
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
      Width           =   5235
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4215
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
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   12435
      _cx             =   21934
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
      FormatString    =   $"FrmAsseetSearch.frx":038A
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
      Left            =   810
      TabIndex        =   12
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
      Left            =   30
      TabIndex        =   13
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Label LblClientName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   195
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3270
      Width           =   795
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
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAssestSearch"
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
Me.DtpDateFromDr.value = ""
Me.DtpDateToDr.value = ""
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

  If frmdriveassestMove.bo = True Then
 frmdriveassestMove.retrive1 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 Else
 frmdriveassestMove.Retrive2 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
End If
End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
    ' Dcombos.GetClientName Me.DCEmp_Name

    Set DCboSearch = New clsDCboSearch
    'Set DCboSearch.Client = Me.DCEmp_Name
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
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    SetDtpickerDate Me.DtpDateFromDr
    SetDtpickerDate Me.DtpDateToDr

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

StrSQL = " SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt, "
StrSQL = StrSQL & "                      dbo.TblEmpAsest.EmpAsID, dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsest.EmpAsestID, dbo.TblEmpAsest.PostedDate, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_ID, dbo.TblEmpAsest.RecordDate,"
StrSQL = StrSQL & "                       dbo.TblEmpAsestDetails.EmpID , dbo.TblEmpAsestDetails.FlagAs"
StrSQL = StrSQL & "  FROM         dbo.TblAssestes INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TblEmpAsestDetails.EmpID = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null)"

    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblEmpAsest.EmpAsID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAsest.EmpAsID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
     

  
        If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAsest.EmpAsID <=" & val(Me.TxtIDTO.text) & ""
       Else
           BolBegine = True
            StrWhere = " Where dbo.TblEmpAsest.EmpAsID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
  
'////////////////////////
 If Me.txtfullcode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.fullcode = '" & Me.txtfullcode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.fullcode = '" & Me.txtfullcode.text & "'"
        End If
    End If
 If Me.TxtFitter.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.TxtFitter.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.TxtFitter.text & "%'"
        End If
    End If
   If txtassest.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAssestes.AsName like '%" & Me.txtassest.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAssestes.AsName like '%" & Me.txtassest.text & "%'"
        End If
    End If
If txtqunatity.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.Qunt like '%" & Me.txtqunatity.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAsestDetails.Qunt like '%" & Me.txtqunatity.text & "%'"
        End If
    End If
    If TxtRemarks.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.Remarks like '%" & Me.TxtRemarks.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAsestDetails.Remarks like '%" & Me.TxtRemarks.text & "%'"
        End If
    End If
  

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAsest.Recorddate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAsest.Recorddate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEmpAsest.Recorddate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEmpAsest.Recorddate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
'''//////////////
 If Not IsNull(Me.DtpDateFromDr.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmpAsest.PostedDate >=" & SQLDate(Me.DtpDateFromDr.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmpAsest.PostedDate >=" & SQLDate(Me.DtpDateFromDr.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateToDr.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblEmpAsest.PostedDate <=" & SQLDate(Me.DtpDateToDr.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblEmpAsest.PostedDate <=" & SQLDate(Me.DtpDateToDr.value, True) & ""
        End If
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblEmpAsest.EmpAsID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
    'Me.lbl(10).Caption = "‰ ÌÃ…«·»Õ÷="
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
    '        Me.lbl(10).Caption = "Search Results=0"
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
              '  Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
              '  Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
            Dim count As Integer
        count = 0
            For i = .FixedRows To .Rows - 1
          If IsNull(rs("FlagAs").value) Then
          count = count + 1
                .TextMatrix(count, .ColIndex("Serial")) = count
              
                .TextMatrix(count, .ColIndex("id")) = IIf(IsNull(rs("EmpAsID").value), "", rs("EmpAsID").value)
                ' .TextMatrix(i, .ColIndex("ID_Aut")) = IIf(IsNull(rs("ID_Aut").value), "", rs("ID_Aut").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(count, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            If Not (IsNull(rs("PostedDate").value)) Then
                    .TextMatrix(count, .ColIndex("DateOp")) = Format(rs("PostedDate").value, "yyyy/M/d")
                End If
                .TextMatrix(count, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(count, .ColIndex("assest")) = IIf(IsNull(rs("AsName").value), "", rs("AsName").value)
                .TextMatrix(count, .ColIndex("Quantity")) = IIf(IsNull(rs("Qunt").value), "", rs("Qunt").value)
              .TextMatrix(count, .ColIndex("remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
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
  Me.Caption = "Search Drive Assest"
lbl(0).Caption = "Assest"
Me.LblClientName.Caption = "Emp Name"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(9).Caption = "From"
lbl(7).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
'lbl(0).Caption = "Total Value"
lbl(8).Caption = "Quantity"
lbl(2).Caption = "Total"
lbl(12).Caption = "Remarks"
'lbl(11).Caption = "Commission"
'Me.lbreg.Caption = "CommisSearch"
Me.lbprocess.Caption = "Order No"
Frame1.Caption = "Date Drive"
'Frame1.Caption = "No. Process"
lbreg.Caption = "Date Order"

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Order"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date Order"
        
        .TextMatrix(0, .ColIndex("DateOp")) = "Date Drive"
       .TextMatrix(0, .ColIndex("assest")) = "Assest"
        .TextMatrix(0, .ColIndex("empname")) = "Emp Name"
         .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
        .TextMatrix(0, .ColIndex("remarks")) = "Remarks"
   
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

