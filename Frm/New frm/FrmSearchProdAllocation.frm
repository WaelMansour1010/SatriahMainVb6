VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearchProdAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ  Œ’Ì’ ŒÿÊÿ «·«‰ «Ã ·«Ê«„— «·‘€·"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "FrmSearchProdAllocation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CboType1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      ItemData        =   "FrmSearchProdAllocation.frx":038A
      Left            =   0
      List            =   "FrmSearchProdAllocation.frx":038C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Õ—ﬂ…"
      Height          =   645
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   5235
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   2460
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·› —Â"
      Height          =   1215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96337923
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96337923
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï  «—ÌŒ"
         Height          =   195
         Index           =   2
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1575
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   5505
      Begin MSDataListLib.DataCombo dcLineID 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483646
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
      Begin MSDataListLib.DataCombo DCranch 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DCStores 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483646
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Œ“‰"
         Height          =   285
         Index           =   12
         Left            =   4065
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   3
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Œÿ «·«‰ «Ã"
         Height          =   285
         Index           =   0
         Left            =   4230
         TabIndex        =   7
         Top             =   1200
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
      FormatString    =   $"FrmSearchProdAllocation.frx":038E
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
      TabIndex        =   2
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
      TabIndex        =   3
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
Attribute VB_Name = "FrmSearchProdAllocation"
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
            DtpDateFrom.value = ""
DtpDateTo.value = ""
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
FrmProductionAllocation.Retrive val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
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
 Dcombos.GetBranches Me.DCranch
 Dcombos.GetLine Me.dcLineID
 Dcombos.GetStores Me.DCStores

    
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
DtpDateTo.value = ""

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

StrSQL = "SELECT     dbo.tblProductionAlloc.ID, dbo.tblProductionAlloc.OPrDate, dbo.tblProductionAlloc.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                       dbo.tblProductionAlloc.StoreId, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.tblProductionAlloc.LineID, dbo.TblProductLine.name,"
StrSQL = StrSQL & "                      dbo.tblProductionAlloc.fromTime, dbo.tblProductionAlloc.toTime, dbo.tblProductionAlloc.NoOfHours, dbo.tblProductionAlloc.TotalSalaries,"
StrSQL = StrSQL & "                      dbo.tblProductionAlloc.TotalElectricals, dbo.tblProductionAlloc.WorkOrderNO, dbo.tblProductionAlloc.totalLineExpenses, dbo.tblProductionAlloc.totalOrderQty,"
StrSQL = StrSQL & "                      dbo.tblProductionAlloc.TotalProductionQty, dbo.tblProductionAlloc.totalMaterialsForItems, dbo.tblProductionAlloc.totalMaterials, dbo.tblProductionAlloc.REMARKS,"
StrSQL = StrSQL & "                      dbo.tblProductionAlloc.UsedPowerPriceH, dbo.tblProductionAlloc.Transaction_ID, dbo.tblProductionAlloc.NoteSerial, dbo.tblProductionAlloc.NoteSerial1,"
StrSQL = StrSQL & "                      dbo.tblProductionAlloc.ReciveDate , dbo.tblProductionAlloc.NoteSerial1V"
StrSQL = StrSQL & " FROM         dbo.tblProductionAlloc LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProductLine ON dbo.tblProductionAlloc.LineID = dbo.TblProductLine.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.tblProductionAlloc.StoreId = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblProductionAlloc.BranchID = dbo.TblBranchesData.branch_id"

    BolBegine = False
    StrWhere = ""

    '///////////////////
        If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.tblProductionAlloc.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblProductionAlloc.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblProductionAlloc.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    

    

    
          If Me.DCranch.text <> "" And (val(DCranch.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.BranchID =" & Me.DCranch.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblProductionAlloc.BranchID =" & Me.DCranch.BoundText & ""
        End If
    End If
 ''//
     If Me.dcLineID.text <> "" And (val(dcLineID.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.LineID =" & Me.dcLineID.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblProductionAlloc.LineID =" & Me.dcLineID.BoundText & ""
        End If
    End If
 ''//
 
 
       If Me.DCStores.text <> "" And (val(DCStores.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.StoreId =" & Me.DCStores.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblProductionAlloc.StoreId =" & Me.DCStores.BoundText & ""
        End If
    End If
     If Not IsNull(Me.DtpDateFrom.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.OPrDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.tblProductionAlloc.OPrDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
                   
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblProductionAlloc.OPrDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.tblProductionAlloc.OPrDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
                   
      End If


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.tblProductionAlloc.id "
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
               
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                        
                If Not (IsNull(rs("OPrDate").value)) Then
                    .TextMatrix(i, .ColIndex("OPrDate")) = Format(rs("OPrDate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
       
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
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
  Me.Caption = "Saerch Production Allocation"

lbprocess.Caption = "No Transection"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lblLW.Caption = "Saerch By"
lbl(3).Caption = "Branch"
lbl(0).Caption = "Production line"
lbl(12).Caption = "Store"
Frame1.Caption = "Priod"
lbl(4).Caption = "From"
lbl(2).Caption = "To"

Frame1.Caption = "Interval"
lbl(4).Caption = "From"
lbl(2).Caption = "To"

 
    
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("OPrDate")) = "RecordDate"
         .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
       .TextMatrix(0, .ColIndex("name")) = "Production line"
    End With
  '
End Sub






