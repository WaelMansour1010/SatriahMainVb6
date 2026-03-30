VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemCostShow 
   Caption         =   "⁄—÷ „ Ê”ÿ «· ﬂ·›… ·’‰›"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   10905
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8040
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10905
      _cx             =   19235
      _cy             =   14182
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmItemCostSow.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   390
         Left            =   15
         TabIndex        =   7
         Top             =   7635
         Visible         =   0   'False
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   945
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   10875
         _cx             =   19182
         _cy             =   1667
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   1
            Left            =   2700
            TabIndex        =   14
            Top             =   30
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
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
         End
         Begin VB.TextBox TxtToInv 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   390
            Visible         =   0   'False
            Width           =   1755
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   405
            Index           =   0
            Left            =   2700
            TabIndex        =   8
            Top             =   390
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            Caption         =   " ‰›Ì–"
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
         Begin MSDataListLib.DataCombo DcboItemName 
            Height          =   315
            Left            =   3300
            TabIndex        =   6
            Top             =   45
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   45
            Width           =   1755
         End
         Begin VB.Label LblCostTransID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ Õ—ﬂ… «· ﬂ·›…"
            Height          =   345
            Index           =   3
            Left            =   1230
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   510
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ Ê”ÿ ”⁄— «· ﬂ·›…"
            Height          =   345
            Index           =   2
            Left            =   1230
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   90
            Width           =   1365
         End
         Begin VB.Label LblLastCost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   30
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›"
            Height          =   315
            Index           =   1
            Left            =   7380
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   45
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   315
            Index           =   0
            Left            =   9930
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   885
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6645
         Left            =   15
         TabIndex        =   1
         Top             =   975
         Width           =   10875
         _cx             =   19182
         _cy             =   11721
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
         BackColorFixed  =   -2147483633
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
         AllowUserResizing=   0
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
         FormatString    =   $"FrmItemCostSow.frx":0081
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   8
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
   End
End
Attribute VB_Name = "FrmItemCostShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cDboSearch As clsDCboSearch

Public Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

            DoAction

        Case 1
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DcboItemName
            FrmItemSearch.Show vbModal
    End Select

End Sub

Public Sub DoAction()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String
    Dim SngQty As Single
    Dim LngItemID As Long
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single, SngTemp4 As Single

    If val(Me.DcboItemName.BoundText) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰› ...!!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DcboItemName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    LngItemID = val(Me.DcboItemName.BoundText)
    StrSQL = "Select * From RptItemTransCus"
    StrSQL = StrSQL + " Where Item_ID=" & LngItemID
    StrSQL = StrSQL + " AND (Transaction_Type=1 OR  Transaction_Type=3)"

    If val(Me.TxtToInv.text) <> 0 Then
        StrSQL = StrSQL + " AND Transaction_ID <" & val(Me.TxtToInv.text)
    End If

    StrSQL = StrSQL + " Order BY Transaction_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FG
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)

                If Not IsNull(rs("Transaction_Date").value) Then
                    .TextMatrix(i, .ColIndex("Transaction_Date")) = DisplayDate(rs("Transaction_Date").value)
                End If
            
                .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                .TextMatrix(i, .ColIndex("TransQty")) = IIf(IsNull(rs("XQty").value), "", rs("XQty").value)
                .TextMatrix(i, .ColIndex("XPrice")) = IIf(IsNull(rs("XPrice").value), "", rs("XPrice").value)
                .TextMatrix(i, .ColIndex("BeforeQty")) = GetItemStockToTrans(LngItemID, val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
                .TextMatrix(i, .ColIndex("BeforeCostPrice")) = val(.TextMatrix(i - 1, .ColIndex("NewCostPrice")))
                '«·„⁄«œ·…
                '(«·ﬂ„Ì… «·»«ﬁÌ… »”⁄— «· ﬂ·›… «·√ŒÌ—  „÷—Ê»« ›Ï  ”⁄— «· ﬂ·›…«·√ŒÌ—)
                '+
                '(«·ﬂ„Ì… «·Ê«—œ… »”⁄— «· ﬂ·›… «·ÃœÌœ „÷—Ê»« ›Ï ”⁄— «· ﬂ·›… «·ÃœÌœ)
                '„ﬁ”„« ⁄·Ï
                '≈Ã„«·Ï «·ﬂ„Ì… «·ÃœÌœ… Ê«·ﬁœÌ„…
            
                SngTemp1 = val(.TextMatrix(i, .ColIndex("BeforeQty"))) * val(.TextMatrix(i, .ColIndex("BeforeCostPrice")))
            
                SngTemp2 = val(.TextMatrix(i, .ColIndex("TransQty"))) * val(.TextMatrix(i, .ColIndex("XPrice")))
            
                SngTemp3 = val(.TextMatrix(i, .ColIndex("BeforeQty"))) + val(.TextMatrix(i, .ColIndex("TransQty")))

                If SngTemp3 <> 0 Then
                    SngTemp4 = (SngTemp1 + SngTemp2) / SngTemp3
                Else
                    SngTemp4 = 0
                End If

                .TextMatrix(i, .ColIndex("NewCostPrice")) = Format(SngTemp4, SystemOptions.SysDefCurrencyForamt)
            
                rs.MoveNext
            Next i

            Me.LblLastCost.Caption = .TextMatrix(.Rows - 1, .ColIndex("NewCostPrice"))
            Me.LblCostTransID.Caption = .TextMatrix(.Rows - 1, .ColIndex("Transaction_ID"))
        Else
        End If

        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub DcboItemName_Change()

    If val(Me.DcboItemName.BoundText) <> 0 Then
        Me.TxtItemCode.text = GetItemCode(Me.DcboItemName.BoundText)
    End If

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DcboItemName
    Set cDboSearch = New clsDCboSearch
    Set cDboSearch.Client = Me.DcboItemName
    Me.Icon = mdifrmmain.ImgLstMenuIcons.ListImages("Num").Picture
    Set Me.Cmd(1).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("BrowseFile").Picture
    Me.Cmd(1).ButtonStyle = impActive

    With Me.FG
        Set GrdBack = New ClsBackGroundPic
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Me.Width = 12000
    Me.Height = 9500
    Resize_Form Me
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Me.DcboItemName.BoundText = GetItemID(Trim(Me.TxtItemCode.text))
    End If

End Sub

