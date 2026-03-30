VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGardSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì«  «·Ã—œ"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   Icon            =   "FrmGardSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   4650
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4545
      _cx             =   8017
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmGardSearch.frx":038A
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
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbEdit 
      Height          =   315
      Left            =   1470
      TabIndex        =   1
      Top             =   2730
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   5
      Top             =   3510
      Width           =   735
      _ExtentX        =   1296
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
      TabIndex        =   6
      Top             =   3510
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   120
      TabIndex        =   7
      Top             =   3510
      Width           =   735
      _ExtentX        =   1296
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   315
      Index           =   5
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3105
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·Ã—œ"
      Height          =   315
      Index           =   6
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2730
      Width           =   1095
   End
End
Attribute VB_Name = "FrmGardSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo  As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                FG.Clear flexClearScrollable, flexClearEverything
                Exit Sub
            End If

            Retrive

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
        FrmGard.Retrive val(FG.TextMatrix(FG.Row, 1))
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("ID")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
                .TextMatrix(Num, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Format((rs("Transaction_Date").value), "yyyy/m/d"))
                .TextMatrix(Num, .ColIndex("Store")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
            End With

            rs.MoveNext
        Next Num

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Set rs = New ADODB.Recordset
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    CenterForm Me

    FormPostion Me, GetPostion
    'XPDtbEdit.Value = Date
    StrSQL = "SELECT * From TblStore"
    fill_combo Me.DCboStoreName, StrSQL
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName
    FG.WallPaper = BG.SearchWallpaper
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
    Set cSearchDcbo = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap
    StrSQL = "select * From QryGardSearch"

    If DCboStoreName.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and StoreName='" & Trim(DCboStoreName.text) & "'"
        Else
            StrWhere = StrWhere + " where StoreName='" & Trim(DCboStoreName.text) & "'"
            Begin = True
        End If
    End If

    If XPDtbEdit.value <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Date=#" & SQLDate(XPDtbEdit.value) & "#"
        Else
            StrWhere = StrWhere + " where Transaction_Date=#" & SQLDate(XPDtbEdit.value) & "#"
            Begin = True
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
        If Not FG.TextMatrix(FG.Row, 1) = "" Then
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

