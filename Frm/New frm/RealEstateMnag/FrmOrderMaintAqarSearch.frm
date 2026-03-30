VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOrderMaintAqarSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·»«  «·’Ì«‰Â"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   Icon            =   "FrmOrderMaintAqarSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1725
      Index           =   1
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3600
      Width           =   5235
      Begin VB.TextBox TxDes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtSearch 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TxtSearchCodeSuper 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpNameSuper 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ê’› «·’Ì«‰Â"
         Height          =   435
         Index           =   8
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ÊŸ› «·’Ì«‰Â"
         Height          =   285
         Index           =   7
         Left            =   4080
         TabIndex        =   26
         Top             =   855
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   255
         Index           =   13
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”ƒÊ· «·’Ì«‰Â"
         Height          =   285
         Index           =   0
         Left            =   4050
         TabIndex        =   22
         Top             =   495
         Width           =   1125
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·ÿ·»"
      Height          =   1035
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   5295
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2490
         TabIndex        =   6
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   515702787
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   2490
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   515702787
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateFromh 
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToh 
         Height          =   315
         Left            =   840
         TabIndex        =   30
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   4680
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
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   1005
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2520
      Width           =   2835
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   360
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
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   1455
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
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8235
      _cx             =   14526
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmOrderMaintAqarSearch.frx":038A
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
      Top             =   5040
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
      Top             =   5040
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
      Top             =   5040
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
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4020
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
      Top             =   4020
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
      Top             =   2640
      Width           =   2775
   End
End
Attribute VB_Name = "FrmOrderMaintAqarSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public m_RetrunType As Integer

Private Sub Cmd_Click(index As Integer)

    Select Case index

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


Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.text = EmpCode
End Sub

Private Sub DtpDateFrom_Change()
 DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
End Sub



Private Sub DtpDateFromH_LostFocus()
        VBA.Calendar = vbCalGreg
           DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub

Private Sub DtpDateTo_Change()
 DtpDateToH.value = ToHijriDate(DtpDateTo.value)
End Sub



Private Sub DtpDateToH_LostFocus()
   VBA.Calendar = vbCalGreg
           DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub fg_Click()

    With Me.FG

        If .row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.row, .ColIndex("id"))) = 0 Then
            Exit Sub
        End If
If m_RetrunType = 1 Then
      FrmLockedOrderMaintenance.TxtOrderNo.text = val(.TextMatrix(.row, .ColIndex("id")))
ElseIf m_RetrunType = 11 Then
      RsExpenses.TXTOrDer_no2.text = val(.TextMatrix(.row, .ColIndex("id")))
      RsExpenses.RetriveOrder
      Else
      
               FrmOrderMaintenance.Retrive val(.TextMatrix(.row, .ColIndex("id")))
               
                
       End If

    End With

End Sub


 

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboEmpNameSuper
  

    Dcombos.GetIqar DcbIqara

    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
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

StrSQL = " SELECT     dbo.TblOrderMaintenance.ID, dbo.TblOrderMaintenance.RecDateH, dbo.TblOrderMaintenance.RecDate, dbo.TblOrderMaintenance.TimOrder, "
 StrSQL = StrSQL & "                     dbo.TblOrderMaintenance.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblOrderMaintenance.EmpID,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Name4, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4, TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2,"
 StrSQL = StrSQL & "                     TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblOrderMaintenance.SuperVM, TblEmployee_1.Emp_Code AS Emp_CodeSup,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Name AS Emp_NameSup, TblEmployee_1.Emp_Name1 AS Emp_Name1Sup, TblEmployee_1.Emp_Name2 AS Emp_Name2Sup,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Name3 AS Emp_Name3Sup, TblEmployee_1.Emp_Name4 AS Emp_Name4Sup, TblEmployee_1.Fullcode AS FullcodeSup,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Namee4 AS Emp_Namee4Sup, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Sup, TblEmployee_1.Emp_Namee2 AS Emp_Namee2Sup,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Namee1 AS Emp_Namee1Sup, TblEmployee_1.Emp_Namee AS Emp_NameeSup, dbo.TblOrderMaintenance.AqrID, dbo.TblAqar.aqarname,"
 StrSQL = StrSQL & "                     dbo.TblOrderMaintenance.LocationIqar, dbo.TblOrderMaintenance.Des, dbo.TblOrderMaintenance.DMY, dbo.TblOrderMaintenance.Cont,"
 StrSQL = StrSQL & "                     dbo.TblOrderMaintenance.EndFateH, dbo.TblOrderMaintenance.EndFate, dbo.TblOrderMaintenance.Lock, dbo.TblOrderMaintenance.LockDateH,"
 StrSQL = StrSQL & "                     dbo.TblOrderMaintenance.LockDate"
StrSQL = StrSQL & " FROM         dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblOrderMaintenance ON TblEmployee_2.Emp_ID = dbo.TblOrderMaintenance.EmpID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblOrderMaintenance.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblOrderMaintenance.BranchID = dbo.TblBranchesData.branch_id"

  StrWhere = ""
    BolBegine = False
   If m_RetrunType = 1 Then
   StrWhere = " Where  dbo.TblOrderMaintenance.Lock <>1"
   BolBegine = True
End If
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.ID >=" & val(Me.TxtIDFrom.text) & ""
       Else
            BolBegine = True
            StrWhere = " Where dbo.TblOrderMaintenance.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
   

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
           BolBegine = True
           StrWhere = " Where dbo.TblOrderMaintenance.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
   
'////////////////////////
 If TxDes.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.Des like '%" & Me.TxDes.text & "%'"
        Else
            BolBegine = True
           StrWhere = " Where dbo.TblOrderMaintenance.Des like '%" & Me.TxDes.text & "%'"
        End If
    End If
  If Me.DcboEmpNameSuper.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.SuperVM=" & Me.DcboEmpNameSuper.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOrderMaintenance.SuperVM=" & Me.DcboEmpNameSuper.BoundText & ""
        End If
    End If
       If Me.DcboEmpName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.EmpID=" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOrderMaintenance.EmpID=" & Me.DcboEmpName.BoundText & ""
        End If
    End If
   If Me.DcbIqara.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid=" & Me.DcbIqara.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblAqar.Aqarid=" & Me.DcbIqara.BoundText & ""
        End If
    End If
   
   


    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.RecDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOrderMaintenance.RecDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblOrderMaintenance.RecDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblOrderMaintenance.RecDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblOrderMaintenance.ID"
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
                    .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", rs("aqarname").value)
                If Not (IsNull(rs("RecDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecDate")) = Format(rs("RecDate").value, "yyyy/M/d")
                End If
             .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
               .TextMatrix(i, .ColIndex("RecDateH")) = IIf(IsNull(rs("RecDateH").value), "", rs("RecDateH").value)
                If SystemOptions.UserInterface = EnglishInterface Then
         .TextMatrix(i, .ColIndex("Emp_NameSup")) = IIf(IsNull(rs("Emp_NameeSup").value), "", rs("Emp_NameeSup").value)
          
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                Else
               .TextMatrix(i, .ColIndex("Emp_NameSup")) = IIf(IsNull(rs("Emp_NameSup").value), "", rs("Emp_NameSup").value)
               .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
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
  Me.Caption = "Search CarRecept"

'Me.LblClientName.Caption = "EngName"
'Me.Label1.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Type"
lbl(8).Caption = "PlateNo"
lbl(2).Caption = "Total"
lbl(9).Caption = "Telephone"

Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
lbl(7).Caption = "Model"
     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("ClientName")) = "ClientName"
        .TextMatrix(0, .ColIndex("Telephone")) = "EngName"
       .TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
       .TextMatrix(0, .ColIndex("mobile")) = "Telephone"
    End With
  '
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub DcbIqara_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

Load FrmAqarSearch
FrmAqarSearch.m_RetrunType = 2020
FrmAqarSearch.show


End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub

Private Sub DcboEmpNameSuper_Change()
DcboEmpNameSuper_Click (0)
End Sub

Private Sub DcboEmpNameSuper_Click(Area As Integer)

   If val(DcboEmpNameSuper.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpNameSuper.BoundText, EmpCode
    TxtSearchCodeSuper.text = EmpCode
End Sub
Private Sub TxtSearchCodeSuper_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCodeSuper.text, EmpID
        DcboEmpNameSuper.BoundText = EmpID
    End If
End Sub
Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

