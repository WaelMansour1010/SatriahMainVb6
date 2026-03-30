VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FixedAssetsSearch1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·ﬁÊ«·»"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "FixedAssetSearch1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   5550
   Begin VB.TextBox TxtCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1290
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2940
      Width           =   2955
   End
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·ﬁ«·» »«·ﬂ«„· ›ﬁÿ"
      Height          =   315
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3345
      Width           =   2535
   End
   Begin VB.TextBox XPTxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1290
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2580
      Width           =   2955
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2505
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5475
      _cx             =   9657
      _cy             =   4419
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FixedAssetSearch1.frx":030A
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
      Left            =   1980
      TabIndex        =   3
      Top             =   3720
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   990
      TabIndex        =   4
      Top             =   3720
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ"
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   2
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3750
      Width           =   2535
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·ﬁ«·»"
      Height          =   315
      Index           =   0
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·ﬁ«·»"
      Height          =   315
      Index           =   1
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2580
      Width           =   1215
   End
End
Attribute VB_Name = "FixedAssetsSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim m_SearchType As Integer

Private m_DcboCustomers As DataCombo

Private m_RetrunType As Integer

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.XPLbl(2).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Else
                Me.XPLbl(2).Caption = "Search Results: " & rs.RecordCount
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
        If Me.RetrunType = 0 Then
    
            FixedAssets.Retrive val(Fg.TextMatrix(Fg.Row, 0))
        ElseIf Me.RetrunType = 1 Then
            FrmItems.DcTemplate.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
            
            ElseIf Me.RetrunType = 2 Then
            frmequipment.DcFixedAssets.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
            
           ElseIf Me.RetrunType = 3 Then
            FrmCars.DcFixedAssets.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
            ElseIf Me.RetrunType = 4 Then
            frmequipment1.DcFixedAssets.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
             
             
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("MemCode")) = IIf(IsNull(rs("code").value), "", (rs("code").value))
                .TextMatrix(Num, .ColIndex("MemNme")) = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
            End With

            rs.MoveNext
        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Activate()

    If Me.SearchType = 1 Then
        Me.Caption = "«·»ÕÀ ⁄‰ «·⁄„·«¡ Ê«·„Ê—œÌ‰"
        Me.XPLbl(1).Caption = "«·ﬂÊœ"
        Me.XPLbl(0).Caption = "«·√”„"
        XPChkSearchType.Caption = "«”„ «·‘Œ’ »«·ﬂ«„·"

        With Me.Fg
            .TextMatrix(0, .ColIndex("MemCode")) = "ﬂÊœ «·„Ê—œ «Ê «·⁄„Ì·"
            .TextMatrix(0, .ColIndex("MemNme")) = "«”„ «·„Ê—œ «Ê «·⁄„Ì·"
        End With

    ElseIf Me.SearchType = 3 Then
        Me.Caption = "«·»ÕÀ ⁄‰ »Ì«‰«  „ﬁ«Ê·Ì «·»«ÿ‰"
        Me.XPLbl(1).Caption = "«·ﬂÊœ"
        Me.XPLbl(0).Caption = "«·√”„"
        XPChkSearchType.Caption = "«”„ «·‘Œ’ »«·ﬂ«„·"

        With Me.Fg
            .TextMatrix(0, .ColIndex("MemCode")) = "ﬂÊœ «·„ ⁄«„·"
            .TextMatrix(0, .ColIndex("MemNme")) = "«”„ «·„ ⁄«„·"
        End With

    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    Dim BG As New ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap

    If Me.SearchType = 0 Then
        StrSQL = "select * From TblEquipments1 "
        Begin = False
    Else
        StrSQL = "select * From TblEquipments1 "
        Begin = False
    End If

    If XPTxtCusID.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and code='" & (XPTxtCusID.text) & "'"
        Else
            StrWhere = StrWhere + " where code='" & (XPTxtCusID.text) & "'"
            Begin = True
        End If
    End If

    If Trim(Me.TxtCustomerName.text) <> "" Then
        If XPChkSearchType.value = Checked Then
            
            If Begin = True Then
                StrWhere = StrWhere + " and Name ='" & Trim(Me.TxtCustomerName.text) & "'"
            Else
                StrWhere = StrWhere + " where Name ='" & Trim(Me.TxtCustomerName.text) & "'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and Name like '%" & Trim(TxtCustomerName.text) & "%'"
            Else
                StrWhere = StrWhere + " where Name like '%" & Trim(TxtCustomerName.text) & "%'"
                Begin = True
            End If
        End If
    End If

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
            Fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Customers Search..."
    XPLbl(1).Caption = "Customer Code"
    XPLbl(0).Caption = "Customer Name"
    XPLbl(2).Caption = "Search Results."
    XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    With Me.Fg
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("MemCode")) = "Customer Code"
        .TextMatrix(0, .ColIndex("MemNme")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Phone")) = "Customer Phone"
    End With

End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue
    'm_SearchType=0 Search For Customers only
    'm_SearchType=1 Search For All table

End Property

Public Property Get DcboCustomers() As DataCombo
    Set DcboCustomers = m_DcboCustomers
End Property

Public Property Set DcboCustomers(ByVal vNewValue As DataCombo)
    Set m_DcboCustomers = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
End Property
