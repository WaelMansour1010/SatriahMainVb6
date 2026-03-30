VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCarAuthoSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ »ÿ«ﬁ… «–‰ «’·«Õ"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "FrmCarAuthoSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— «·‘€·"
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3360
      Width           =   3795
      Begin VB.TextBox TxtWorkOrderTo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtWorkOrder 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   9
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   7
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   24
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
      TabIndex        =   19
      Top             =   4440
      Width           =   6675
      Begin VB.TextBox TxtPlateNo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox TxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   195
         Index           =   8
         Left            =   2460
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰ «·⁄„Ì·"
         Height          =   195
         Index           =   0
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2880
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
         Format          =   84475907
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
         Format          =   84475907
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
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6675
      _cx             =   11774
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
      FormatString    =   $"FrmCarAuthoSearch.frx":038A
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
   Begin MSDataListLib.DataCombo DCEmp_Name 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   4110
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "DCEmp_Name"
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   12
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
      TabIndex        =   13
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
      TabIndex        =   14
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label LblClientName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   195
      Left            =   5835
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4110
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmCarAuthoSearch"
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
        If FrmBillCarMaintExtra.Ch = True Then
          GetData1
Else
 GetData
            End If
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


Private Sub fg_Click()

   If FrmBillCarMaintExtra.Ch = True Then
 FrmBillCarMaintExtra.retrive1 (val(Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id"))))
 Else
  FrmCarAuthontication.Retrive2 (val(Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id"))))
  FrmCarAuthontication.TxtAmoutAccept.Text = 0
    FrmCarAuthontication.TxtFirstPrice.Text = 0
    FrmCarAuthontication.TXtCarMeter.Text = ""
    FrmCarAuthontication.DcbOrderStatus.ListIndex = 0
FrmCarAuthontication.ComGranty.ListIndex = 2
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
     Dcombos.GetClientName Me.DCEmp_Name
      If SystemOptions.UserInterface = EnglishInterface Then
     
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"

             Else
  
 DcbOrderStatus.AddItem "ÃœÌœ"
DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


    End If
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.fg
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


StrSQL = "SELECT * FROM TblCardAuthorizationReform "
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
 
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    '///////////////////
     If TxtPhone.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Telephone =" & Me.TxtPhone.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.Telephone =" & Me.TxtPhone.Text & ""
        End If
    End If
'////////////////////////
 If TxtPlateNo.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.PlateNo like '%" & Me.TxtPlateNo.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.PlateNo like '%" & Me.TxtPlateNo.Text & "%'"
        End If
    End If
   If Me.DCEmp_Name.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID=" & Me.DCEmp_Name.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.ID=" & Me.DCEmp_Name.BoundText & ""
        End If
    End If

 
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------
End If
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReform.ID"
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

        With Me.fg
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
                                If rs("OrderStatus").value <> 2 Then
                                Exit Sub
                                End If
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                .TextMatrix(i, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
              '  .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub
Public Sub GetData1()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "SELECT * FROM TblCardAuthorizationReform where OrderStatus = 2 and (WorkOrder )<> 0 and(NOT (WorkOrder IS NULL)) "
    'BolBegine = True
    'StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
            StrWhere = StrWhere & " and  ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    
    If val(Me.TxtIDTO.Text) <> 0 Then
            StrWhere = StrWhere & " AND ID <=" & val(Me.TxtIDTO.Text) & ""
    End If
    
      If val(Me.TxtWorkOrder.Text) <> 0 Then
            StrWhere = StrWhere & " and WorkOrder >=" & val(Me.TxtWorkOrder.Text) & ""
        End If
    
    If val(Me.TxtWorkOrderTo.Text) <> 0 Then
            StrWhere = StrWhere & " AND WorkOrder <=" & val(Me.TxtWorkOrderTo.Text) & ""
    End If
    
    '///////////////////
     If TxtPhone.Text <> "" Then
            StrWhere = StrWhere & " AND Telephone =" & Me.TxtPhone.Text & ""
    End If
'////////////////////////
 If TxtPlateNo.Text <> "" Then
            StrWhere = StrWhere & " AND PlateNo like '%" & Me.TxtPlateNo.Text & "%'"
        End If
    
   If Me.DCEmp_Name.Text <> "" And val(Me.DCEmp_Name.BoundText) <> 0 Then
            StrWhere = StrWhere & " AND ID=" & Me.DCEmp_Name.BoundText & ""
    End If


    If Not IsNull(Me.DtpDateFrom.value) Then
            StrWhere = StrWhere & " AND RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By ID"
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

        With Me.fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
   Dim k As Integer
  
            For i = .FixedRows To .Rows - 1
           
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("WorkOrder")) = IIf(IsNull(rs("WorkOrder").value), "", rs("WorkOrder").value)
                    
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
          
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                .TextMatrix(i, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
              '  .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
             
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

'    End If
End If
End Sub
Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Search CardAuthorizationReform"

Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Telephone"
lbl(8).Caption = "PlateNo"
lbl(2).Caption = "Total"
Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
     With Me.fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("ClientName")) = "ClientName"
        .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
       .TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
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

