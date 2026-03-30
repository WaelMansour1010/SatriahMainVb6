VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmStoreSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·„Œ«“‰"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   Icon            =   "FrmStoreSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2580
      Width           =   5115
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label LblClientName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·›—⁄"
         Height          =   195
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   645
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
      Height          =   1125
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   11715
      Begin VB.TextBox TxtAdres 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox TxtEmpStpre 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5700
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   1995
      End
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   4515
      End
      Begin VB.TextBox TxtCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1875
      End
      Begin VB.TextBox Txtxtelephone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5700
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox TxtSotreNme 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   4515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰Ê«‰ «·„Œ“‰"
         Height          =   195
         Index           =   9
         Left            =   10635
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«„Ì‰ «·„” Êœ⁄"
         Height          =   195
         Index           =   7
         Left            =   7755
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„·«ÕŸ« "
         Height          =   195
         Index           =   4
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„Œ“‰"
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
         Caption         =   "Â« › «·„Œ“‰"
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
         Caption         =   "«”„ «·„Œ“‰"
         Height          =   195
         Index           =   0
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1020
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
      Left            =   1650
      TabIndex        =   5
      Top             =   4440
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
      Top             =   4440
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
      Top             =   4440
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
      TabIndex        =   27
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmStoreSearch.frx":038A
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
Attribute VB_Name = "FrmStoreSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public mIndex As Long
Public RetrunFrm As Form

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


Private Sub Fg_Click()

If mIndex = 1 Then
    
     'FrmEditUsers.ListStoreSelected.AddItem ListStoreall.List(i)
        'FrmEditUsers.ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(i)
        
        Dim i As Long

        For i = 0 To FrmEditUsers.ListStoreall.ListCount - 1
            If FrmEditUsers.ListStoreall.ItemData(i) = val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id"))) Then
                FrmEditUsers.ListStoreall.Selected(i) = True
                
                
            End If
        Next

        'FrmEditUsers.ListStoreall.Selected(val(Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id")))) = True
Else
    FrmStoreData.Retrive (val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id"))))
End If


End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
 My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    fill_combo dcBranch, My_SQL
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
    ' Dcombos.GetClientName Me.DCEmp_Name
'      If SystemOptions.UserInterface = EnglishInterface Then
     
'        Me.DcbOrderStatus.AddItem "New"
'        Me.DcbOrderStatus.AddItem "Accept Customer"
'        Me.DcbOrderStatus.AddItem "Final Maintenance"

'             Else
  
' DcbOrderStatus.AddItem "ÃœÌœ"
'DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
'DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


'    End If
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

    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 '   SetDtpickerDate Me.DtpDateFrom
 '   SetDtpickerDate Me.DtpDateTo

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

    StrSQL = " SELECT     dbo.TblStore.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.Remarks, dbo.TblStore.Account_Code, "
   StrSQL = StrSQL & "                   dbo.TblStore.Account_Code1, dbo.TblStore.Account_Code2, dbo.TblStore.Account_Code3, dbo.TblStore.linked, dbo.TblStore.Code, dbo.TblStore.StoreNamee,"
  StrSQL = StrSQL & "                    dbo.TblStore.ParetnAccount, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_id , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblStore.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblStore.Emp_ID = dbo.TblEmployee.Emp_ID"
'StrSQL = "SELECT * FROM TblCardAuthorizationReform "
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblStore.StoreID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.StoreID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
 

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.StoreID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.StoreID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    '///////////////////
     If Txtxtelephone.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.StorePhone '%" & Me.Txtxtelephone.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.StorePhone like '%" & Me.Txtxtelephone.Text & "%'"
        End If
    End If
'////////////////////////
 If TxtSotreNme.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.StoreName like '%" & Me.TxtSotreNme.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.StoreName like '%" & Me.TxtSotreNme.Text & "%'"
        End If
    End If
    ''/////////////
     If TxtAdres.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.StoreAdress like '%" & Me.TxtAdres.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.StoreAdress like '%" & Me.TxtAdres.Text & "%'"
        End If
    End If
    '''./////////
     If TXTCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.Code like '%" & Me.TXTCode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.Code like '%" & Me.TXTCode.Text & "%'"
        End If
    End If
    '''///////////////
     If TxtEmpStpre.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.TxtEmpStpre.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.TxtEmpStpre.Text & "%'"
        End If
    End If
    ''///////////
    If TxtRemarks.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.Remarks like '%" & Me.TxtRemarks.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.Remarks like '%" & Me.TxtRemarks.Text & "%'"
        End If
    End If
   If Me.dcBranch.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblStore.BranchId=" & Me.dcBranch.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblStore.BranchId=" & Me.dcBranch.BoundText & ""
        End If
    End If

 

  '  If Not IsNull(Me.DtpDateFrom.value) Then
  '      If BolBegine = True Then
  '          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
  '      Else
  '          BolBegine = True
  ''          StrWhere = " Where dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
   '     End If
   ' End If

   ' If Not IsNull(Me.DtpDateTo.value) Then
   '     If BolBegine = True Then
   ''         StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    ''    Else
     '       BolBegine = True
     '       StrWhere = " Where  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     '   End If
    'End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblStore.StoreID "
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

        With Me.FG
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
           '     If rs("OrderStatus").value <> 2 Then
           '     Exit Sub
           '     End If
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
                        
               ' If Not (IsNull(rs("RecordDate").value)) Then
               '     .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
               ' End If
             .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Code").value), "", rs("Code").value)
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(i, .ColIndex("address")) = IIf(IsNull(rs("StoreAdress").value), "", rs("StoreAdress").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("StorePhone").value), "", rs("StorePhone").value)
               .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                .TextMatrix(i, .ColIndex("branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
              '  .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
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
  Me.Caption = "Search CardAuthorizationReform"

Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "Remarks"
lbl(3).Caption = "SotreCode"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "StoreName"
lbl(8).Caption = "Telephone"
lbl(2).Caption = "Total"
lbl(9).Caption = "Address"
lbl(7).Caption = "Storekeeper"
'Me.lbreg.Caption = "Date Registration"
Fra(0).Caption = "By"
Me.lbprocess.Caption = "Process No"
     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Proce"
        .TextMatrix(0, .ColIndex("branch")) = "Branch Name"
         .TextMatrix(0, .ColIndex("code")) = "StoteCode"
        .TextMatrix(0, .ColIndex("ClientName")) = "StoteName"
       .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
       .TextMatrix(0, .ColIndex("address")) = "Address "
         .TextMatrix(0, .ColIndex("empname")) = "Storekeeper"
        .TextMatrix(0, .ColIndex("remark")) = "Remarks"
       '.TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
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

