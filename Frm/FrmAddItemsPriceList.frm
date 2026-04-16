VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddItemsPriceList 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈÷«›… √’‰«› ≈·Ï ﬁ«∆„… «·√”⁄«—"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
   Icon            =   "FrmAddItemsPriceList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ImpulseButton.ISButton Cmd 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   420
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      Caption         =   "..."
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
      ColorButton     =   14871017
   End
   Begin VB.TextBox TxtSupID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   630
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2130
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox XPTxtPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   2760
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   1575
   End
   Begin ImpulseButton.ISButton XPBtnPass 
      Height          =   330
      Left            =   2250
      TabIndex        =   2
      Top             =   795
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmAddItemsPriceList.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DCboItemsName 
      Height          =   315
      Left            =   570
      TabIndex        =   0
      Top             =   420
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   3345
      Left            =   120
      TabIndex        =   3
      Top             =   1170
      Width           =   5175
      _cx             =   9128
      _cy             =   5900
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmAddItemsPriceList.frx":0924
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
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   375
      Left            =   900
      TabIndex        =   5
      Top             =   5040
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmAddItemsPriceList.frx":097B
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   5040
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmAddItemsPriceList.frx":0D15
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnRemove 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   4560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmAddItemsPriceList.frx":10AF
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      LowerToggledContent=   0   'False
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5340
      X2              =   0
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label XPLblSupName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   60
      Width           =   4185
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”⁄—"
      Height          =   315
      Index           =   2
      Left            =   4380
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   810
      Width           =   915
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   1
      Left            =   4380
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ê—œ"
      Height          =   315
      Index           =   0
      Left            =   4380
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "FrmAddItemsPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo  As clsDCboSearch

Private Sub Cmd_Click()
    ShowDialogItemsSearch Me.DCboItemsName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyF3 Then
        XPBtnRemove_Click
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim StrSQL As String
    Dim StrList As String
    Dim rs As ADODB.Recordset
    Dim RsItems As ADODB.Recordset
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion

    Set Me.Cmd.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("BrowseFile").Picture
    Cmd.ButtonPositionImage = impRightOfText
    Cmd.ButtonStyle = impActive

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DCboItemsName
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboItemsName

    Set GrdBack = New ClsBackGroundPic
    Set Me.FG.WallPaper = GrdBack.Picture
    StrSQL = "select * From TblItems"
    Set RsItems = New ADODB.Recordset
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    StrList = FG.BuildComboList(RsItems, "ItemName", "ItemID")

    If StrList <> "" Then
        FG.ColComboList(FG.ColIndex("Item")) = StrList
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim Msg As String
    Dim RsPassData As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    RsTemp.Open "CusJuncItem", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For RowNum = 1 To FG.Rows - 2
        Set RsTest = New ADODB.Recordset
        StrSQL = "select * From CusJuncItem where CusID=" & Trim(TxtSupID.text)
        StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Item"))
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            Msg = " „  ”ÃÌ· «·’‰›" & Chr(13)
            Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Item")) & Chr(13)
            Msg = Msg + "„⁄ Â–« «·„Ê—œ „‰ ﬁ»·"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        RsTest.Close
    Next RowNum

    For RowNum = 1 To FG.Rows - 2
        RsTemp.AddNew
        RsTemp("ID").value = new_id("CusJuncItem", "ID", "", True)
        RsTemp("CusID").value = Trim(TxtSupID.text)
        RsTemp("ItemID").value = FG.TextMatrix(RowNum, FG.ColIndex("Item"))
        RsTemp("ItemPrice").value = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
        RsTemp("LastUpdate").value = Date
        RsTemp.update

        With FrmMainPriceList.FgMain

            If .TextMatrix(.Rows - 1, .ColIndex("Tree")) <> "" Then
                .Rows = .Rows + 1
            End If

            Set RsPassData = New ADODB.Recordset
            StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Item"))
            RsPassData.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsPassData.EOF Or RsPassData.BOF) Then
                .Rowdata(.Rows - 1) = RsTemp("ID").value
                .TextMatrix(.Rows - 1, .ColIndex("DefalutPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
                .TextMatrix(.Rows - 1, .ColIndex("Tree")) = IIf(IsNull(RsPassData("ItemName").value), "", RsPassData("ItemName").value)
                .TextMatrix(.Rows - 1, .ColIndex("ItemID")) = IIf(IsNull(RsPassData("ItemID").value), "", RsPassData("ItemID").value)
                .TextMatrix(.Rows - 1, .ColIndex("ItemCode")) = IIf(IsNull(RsPassData("ItemCode").value), "", RsPassData("ItemCode").value)
                .TextMatrix(.Rows - 1, .ColIndex("LastUpdate")) = Format(Date, "YYYY/MM/DD")
            End If

            RsPassData.Close
        End With

    Next RowNum

    FrmMainPriceList.GetSupPriceSetting
    RsTemp.Close
    Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnPass_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim LngSearch As Long

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·’‰› √Ê·«" & Chr(13)
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If XPTxtPrice.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ ”⁄— «·’‰› ﬁ»· ≈÷«› Â" & Chr(13)
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(XPTxtPrice.text) Then
        Msg = "”⁄— «·’‰› ·«»œ √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…" & Chr(13)
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    With FG
        LngSearch = .FindRow(DCboItemsName.BoundText, , .ColIndex("Item"))

        If LngSearch > -1 Then
            Msg = "·ﬁœ  „  ”ÃÌ· Â–« «·’‰› „‰ ﬁ»·" & Chr(13)
            Msg = Msg + "·« Ì„ﬂ‰  ”ÃÌ· «·’‰› ·‰›” «·„Ê—œ √ﬂÀ— „‰ „—…" & Chr(13)
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        .TextMatrix(.Rows - 1, .ColIndex("Item")) = DCboItemsName.BoundText
        .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text
        .Rows = .Rows + 1
        .AutoSize 0, .Cols - 1, False
        DCboItemsName.text = ""
        XPTxtPrice.text = ""
    End With

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.Rows > 1 Then
        If FG.Rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
        Else

            If FG.Rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Row)
                End If
            End If
        End If
    End If

    Exit Sub
ErrTrap:
 
End Sub

Private Sub XPTxtPrice_KeyDown(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyReturn Then
        XPBtnPass_Click
        DCboItemsName.SetFocus
    End If

End Sub

Private Sub XPTxtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtPrice.text, 0)
End Sub

