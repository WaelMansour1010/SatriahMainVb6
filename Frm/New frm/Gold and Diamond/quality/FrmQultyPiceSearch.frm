VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmQultyPiceSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÃÊœ… «·ﬁÿ⁄"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "FrmQultyPiceSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   8085
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2760
      Width           =   3495
      Begin VB.TextBox TxtCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂÊœ"
         Height          =   195
         Index           =   7
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Index           =   4
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1725
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3480
      Width           =   7995
      Begin VB.TextBox TxtRemark 
         Alignment       =   1  'Right Justify
         Height          =   585
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   900
         Width           =   6675
      End
      Begin VB.TextBox TxtNameE 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   3795
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   3795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ« "
         Height          =   195
         Index           =   3
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ ≈‰Ã·Ì“Ì"
         Height          =   195
         Index           =   8
         Left            =   6420
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ ⁄—»Ì"
         Height          =   195
         Index           =   0
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   735
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
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
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      _cx             =   14102
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
      FormatString    =   $"FrmQultyPiceSearch.frx":038A
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
      Left            =   1770
      TabIndex        =   6
      Top             =   5280
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
      Left            =   930
      TabIndex        =   7
      Top             =   5280
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
      Left            =   150
      TabIndex        =   8
      Top             =   5280
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
      Left            =   8640
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
      TabIndex        =   9
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmQultyPiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public calltype As Integer

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
    Select Case Index

        Case 0
    
 GetData
           
        Case 1
            clear_all Me
 

        Case 2
            Unload Me
    End Select

End Sub




Private Sub Fg_Click()

 On Error GoTo ErrTrap
If Me.calltype = 0 Then
    
   
 
    FrmQualityPices.FindRec val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))


ElseIf Me.calltype = 1 Then
FrmItems.TxtPartNo.text = Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("code"))

End If
Unload Me
ErrTrap:

End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
     'Dcombos.GetClientName Me.DCEmp_Name
    '  If SystemOptions.UserInterface = EnglishInterface Then
     
    '  Me.DcbOrderStatus.AddItem "New"
    '    Me.DcbOrderStatus.AddItem "Accept Customer"
    '    Me.DcbOrderStatus.AddItem "Final Maintenance"

     '        Else
  
' DcbOrderStatus.AddItem "ÃœÌœ"
'DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
'DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


'    End If
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

   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
StrSQL = "SELECT * FROM TblQuPices "
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " id >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where id >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    ' If val(FrmCarAuthontication.TxtOrder.text) <> 0 Then
    '    If BolBegine = True Then
    '        StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ID =" & val(FrmCarAuthontication.TxtOrder.text) & ""
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblCardAuthorizationReform.ID =" & val(FrmCarAuthontication.TxtOrder.text) & ""
    '    End If
    'End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND id <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where id <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
     If TxtCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND code  like '%" & Me.TxtCode.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where code  like '%" & Me.TxtCode.text & "%'"
        End If
    End If
'////////////////////////
 If Me.TxtName.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND name like '%" & Me.TxtName.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where name like '%" & Me.TxtName.text & "%'"
        End If
    End If
    If Me.TxtNameE.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND nameE like '%" & Me.TxtNameE.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where nameE like '%" & Me.TxtNameE.text & "%'"
        End If
    End If
        If Me.TxtRemark.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Remrks like '%" & Me.TxtRemark.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where Remrks like '%" & Me.TxtRemark.text & "%'"
        End If
    End If
  ' If Me.DCEmp_Name.BoundText <> "" Then
  '      If BolBegine = True Then
  ''          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID=" & Me.DCEmp_Name.BoundText & ""
   '     Else
   '         BolBegine = True
   '         StrWhere = " Where dbo.TblCardAuthorizationReform.ID=" & Me.DCEmp_Name.BoundText & ""
   '     End If
   ' End If

   ' If Me.DCUser.BoundText <> "" Then
   ''     If BolBegine = True Then
    ''        StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
     ''   Else
      ''      BolBegine = True
       '     StrWhere = " Where    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
       ' End If
    'End If

   ' If Not IsNull(Me.DtpDateFrom.value) Then
   '     If BolBegine = True Then
   '         StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
   '     Else
   '         BolBegine = True
   '         StrWhere = " Where dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
   '     End If
   ' End If

   ' If Not IsNull(Me.DtpDateTo.value) Then
   '     If BolBegine = True Then
   '         StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
   '     Else
   '         BolBegine = True
   '         StrWhere = " Where  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
   '     End If
   ' End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By id "
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
               ' If rs("OrderStatus").value <> 2 Then
               ' Exit Sub
               ' End If
                
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                 .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
               ' If Not (IsNull(rs("RecordDate").value)) Then
               '     .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
               ' End If
            
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs("nameE").value), "", rs("nameE").value)
                .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(rs("Remrks").value), "", rs("Remrks").value)
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
    
  Me.Caption = "Search ÛQuality Pices"

'Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "Code"
lbl(7).Caption = "Code"
lbl(3).Caption = "Remarks"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Name Arb"
lbl(8).Caption = " Name ENG"
lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("code")) = "Code"
       ' .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("name")) = " Name Arb"
        .TextMatrix(0, .ColIndex("namee")) = " Name ENG"
       .TextMatrix(0, .ColIndex("remark")) = "Remarks"
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

