VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBalanceSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "FrmBalanceSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   5880
   Begin VB.Frame Fra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "⁄Ê«„· »ÕÀ ≈÷«›Ì…"
      ForeColor       =   &H00000080&
      Height          =   1395
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4380
      Width           =   5805
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   885
      End
      Begin VB.TextBox TxtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtItemSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   3075
      End
      Begin VB.CheckBox ChkSerialSearchType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "»ÕÀ „ÿ«»ﬁ"
         Height          =   285
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCboItem 
         Height          =   315
         Left            =   540
         TabIndex        =   6
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdItemSearch 
         Height          =   345
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBalanceSearch.frx":038A
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "«”„ «·’‰›"
         Height          =   315
         Index           =   8
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂ„Ì… «·’‰›"
         Height          =   315
         Index           =   7
         Left            =   2670
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”Ì—Ì«· «·’‰›"
         Height          =   315
         Index           =   4
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”⁄— «·’‰›"
         Height          =   315
         Index           =   5
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   615
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   345
         Index           =   6
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   885
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ· ›Ï «·› —…"
      Height          =   1065
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2820
      Width           =   2085
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   96075777
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   96075777
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   3
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2820
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5835
      _cx             =   10292
      _cy             =   4842
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
      FormatString    =   $"FrmBalanceSearch.frx":0924
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
      Left            =   2190
      TabIndex        =   2
      Top             =   3195
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   12
      Top             =   3930
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
      Left            =   840
      TabIndex        =   13
      Top             =   3930
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
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3930
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   4470
      TabIndex        =   5
      Top             =   3930
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ „ ﬁœ„..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmBalanceSearch.frx":09E5
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmBalanceSearch.frx":0D7F
      ColorToggledHoverText=   192
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   9
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3540
      Width           =   2325
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   315
      Index           =   2
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2850
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   285
      Index           =   1
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3210
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBalanceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cSearchDcbo(1)  As clsDCboSearch
Public mTransaction_Type As Integer
Public mIndex As Integer
Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                Me.lbl(9).Caption = "‰ ÌÃ… «·»ÕÀ : ’›—"
                Fg.Clear flexClearScrollable, flexClearEverything
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Me.lbl(9).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Retrive

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
            DtpFrom.value = Date
            DtpTo.value = Date
            DtpFrom.value = Null
            DtpTo.value = Null

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DCboItem
    FrmItemSearch.show vbModal
End Sub

Private Sub CmdShowMoreOptions_Click()

    If CmdShowMoreOptions.value = True Then
        Me.Fra(1).Visible = True
        Me.Height = Me.Fra(1).top + Fra(1).Height + 400
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Fra(1).top + 400
    
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
    
    If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
        If mTransaction_Type = 30 And mIndex = 0 Then
            FrmNewGard.Retrive val(Fg.TextMatrix(Fg.Row, 1))
        ElseIf mTransaction_Type = 30 And mIndex = 1 Then
            FrmNewGard1.Retrive val(Fg.TextMatrix(Fg.Row, 1))
        Else
        
            FrmOpeningBalance.Retrive val(Fg.TextMatrix(Fg.Row, 1))
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
                .TextMatrix(Num, .ColIndex("ID")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
                .TextMatrix(Num, .ColIndex("Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
                .TextMatrix(Num, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Format((rs("Transaction_Date").value), "yyyy/m/d"))
                .TextMatrix(Num, .ColIndex("Store")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
            End With

            rs.MoveNext
        Next Num

    End If

    Fg.SetFocus
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
    Dim Dcombos As ClsDataCombos

    Set rs = New ADODB.Recordset

    CenterForm Me

    FormPostion Me, GetPostion
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    SetDtpickerDate Me.DtpFrom
    SetDtpickerDate Me.DtpTo

    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetItemsNames Me.DCboItem
    Fg.WallPaper = BG.SearchWallpaper
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DCboStoreName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboItem
    CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap
    'StrSQL = "select * From QryOpenBalanceSearch"
    StrSQL = " SELECT Distinct Transactions.Transaction_ID, Transactions.Transaction_Date," & "Transactions.Transaction_Type, Transactions.StoreID, TblStore.StoreName," & "Transactions.Transaction_Serial "
    StrSQL = StrSQL + " FROM (TblStore INNER JOIN Transactions ON TblStore.StoreID =" & "Transactions.StoreID) INNER JOIN (TblItems INNER JOIN Transaction_Details ON " & "TblItems.ItemID = Transaction_Details.Item_ID) ON Transactions.Transaction_ID = " & "Transaction_Details.Transaction_ID "
     
    If mTransaction_Type = 0 Then
        StrWhere = " Where(Transactions.Transaction_Type=3) "
    Else
         StrWhere = " Where(Transactions.Transaction_Type=" & mTransaction_Type & " ) "
    End If
    StrWhere = StrWhere & "  AND      Transactions.BranchId in(" & Current_branchSql & ")"
    Begin = True

    If Trim(Me.TxtSerial.Text) <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Serial Like '" & Trim(TxtSerial.Text) & "'"
        Else
            StrWhere = StrWhere + " where Transaction_Serial Like '" & Trim(TxtSerial.Text) & "'"
            Begin = True
        End If
    End If

    If DCboStoreName.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.StoreID=" & Trim(DCboStoreName.BoundText) & ""
        Else
            StrWhere = StrWhere + " where Transactions.StoreID=" & Trim(DCboStoreName.BoundText) & ""
            Begin = True
        End If
    End If

    If Not IsNull(Me.DtpFrom.value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Date >=" & SQLDate(DtpFrom.value, True) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Date >=" & SQLDate(DtpFrom.value, True) & ""
            Begin = True
        End If
    End If

    If Not IsNull(Me.DtpTo.value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Date <=" & SQLDate(DtpTo.value, True) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Date <=" & SQLDate(DtpTo.value, True) & ""
            Begin = True
        End If
    End If

    If Me.DCboItem.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
            Begin = True
        End If
    End If

    If val(TxtItemQty.Text) > 0 Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Quantity=" & val(TxtItemQty.Text) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Quantity=" & val(TxtItemQty.Text) & ""
            Begin = True
        End If
    End If

    If val(TxtItemPrice.Text) > 0 Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Price=" & val(TxtItemPrice.Text) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Price=" & val(TxtItemPrice.Text) & ""
            Begin = True
        End If
    End If

    If Trim(Me.TxtItemSerial.Text) <> "" Then
        If ChkSerialSearchType.value = vbChecked Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.Text) & "'"
            Else
                StrWhere = StrWhere + " where Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.Text) & "'"
                Begin = True
            End If

        ElseIf ChkSerialSearchType.value = vbUnchecked Then

            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.Text) & "%'"
            Else
                StrWhere = StrWhere + " where Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.Text) & "%'"
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
        If Fg.TextMatrix(Fg.Row, Fg.ColIndex("ID")) <> "" And Me.ActiveControl Is Fg Then
            Fg_Click
        ElseIf Shift = vbCtrlMask Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.Text) <> "" Then
            Me.DCboItem.BoundText = GetItemID(Me.TxtItemCode.Text)
        End If
    End If

End Sub

