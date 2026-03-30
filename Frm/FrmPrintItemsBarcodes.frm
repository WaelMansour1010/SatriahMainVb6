VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPrintItemsBarcodes 
   Caption         =   "ÿ»«⁄… «ﬂÊ«œ «·√’‰«›"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   Icon            =   "FrmPrintItemsBarcodes.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8670
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   6450
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8670
      _cx             =   15293
      _cy             =   11377
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPrintItemsBarcodes.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.TextBox TxtQty 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   1620
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   405
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   714
         ButtonStyle     =   1
         ButtonPositionImage=   1
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
         ButtonImage     =   "FrmPrintItemsBarcodes.frx":0418
         DrawFocusRectangle=   0   'False
      End
      Begin VB.ComboBox CboActions 
         Height          =   315
         Left            =   4470
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   2760
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1170
         Index           =   0
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5250
         Width           =   8610
         _cx             =   15187
         _cy             =   2064
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1095
            Index           =   1
            Left            =   2550
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   30
            Width           =   6015
            _cx             =   10610
            _cy             =   1931
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
            Caption         =   "«·»ÕÀ ⁄‰ ’‰›"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   1
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
            Begin VB.CheckBox XPChkSearchType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÿ«»ﬁ… Õ«·… «·√Õ—› "
               Height          =   285
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   240
               Width           =   1965
            End
            Begin VB.CheckBox XPChkFullMuch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ·„… »«·ﬂ«„·"
               Height          =   285
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   630
               Width           =   1335
            End
            Begin VB.TextBox XPTxtSearchValue 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2490
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   600
               Width           =   3390
            End
            Begin VB.ComboBox XPCboSearchType 
               Height          =   315
               Left            =   2490
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   255
               Width           =   3390
            End
            Begin ImpulseButton.ISButton XPBtnSearch 
               Height          =   405
               Left            =   120
               TabIndex        =   18
               Top             =   570
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   714
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
               ButtonImage     =   "FrmPrintItemsBarcodes.frx":07B2
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   465
            Left            =   90
            TabIndex        =   10
            Top             =   570
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌœ Œ’«∆’ «·ÿ»«⁄…"
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
            ButtonImage     =   "FrmPrintItemsBarcodes.frx":0B4C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   4
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   270
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«› «·„Õœœ…"
            Height          =   285
            Index           =   3
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   300
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«› "
            Height          =   285
            Index           =   1
            Left            =   1530
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   60
            Width           =   945
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   4785
         Left            =   30
         TabIndex        =   1
         Top             =   450
         Width           =   8610
         _cx             =   15187
         _cy             =   8440
         Appearance      =   2
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPrintItemsBarcodes.frx":0EE6
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
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
         Caption         =   "⁄œœ «·√” ﬂÌ—«  ·ﬂ· ’‰›"
         Height          =   405
         Index           =   5
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ÿ—Ìﬁ… «· ÕœÌœ"
         Height          =   405
         Index           =   0
         Left            =   7245
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   30
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmPrintItemsBarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadData()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    StrSQL = "SELECT TblItems.ItemID, TblItems.ItemCode, TblItems.ItemName," & "Groups.GroupName,TblItems.HaveSerial,TblItems.SallingPrice "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID;"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg

        If Not (rs.BOF Or rs.EOF) Then
            Me.Lbl(2).Caption = rs.RecordCount
            .Rows = .FixedRows
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(i, .ColIndex("SallingPrice")) = IIf(IsNull(rs("SallingPrice").value), "", rs("SallingPrice").value)
                .TextMatrix(i, .ColIndex("Qty")) = 1

                If rs("HaveSerial").value = True Then
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexUnchecked
                End If

                rs.MoveNext
            Next i

        Else
            Me.Lbl(2).Caption = 0
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing
End Sub

Private Sub CboActions_Change()
    Cmd_Click
End Sub

Private Sub CboActions_Click()
    Cmd_Click
End Sub

Private Sub Cmd_Click()
    Dim Msg As String
    Dim i As Integer
    On Error GoTo ErrTrap

    Select Case Me.CboActions.ListIndex

        Case -1
            Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «· ÕœÌœ...!!"
            Me.CboActions.SetFocus
            SendKeys "F4"
            Exit Sub

        Case 0
            Fg.Cell(flexcpChecked, Fg.FixedRows, Fg.ColIndex("Print"), Fg.Rows - 1, Fg.ColIndex("Print")) = flexChecked
        
        Case 1
            Fg.Cell(flexcpChecked, Fg.FixedRows, Fg.ColIndex("Print"), Fg.Rows - 1, Fg.ColIndex("Print")) = flexUnchecked

        Case 2

            For i = 1 To Fg.Rows - 1
                Fg.Cell(flexcpChecked, i, Fg.ColIndex("Print"), i, Fg.ColIndex("Print")) = Fg.Cell(flexcpChecked, i, Fg.ColIndex("HaveSerial"), i, Fg.ColIndex("HaveSerial"))
            Next i

        Case 3

            For i = 1 To Fg.Rows - 1

                If Fg.Cell(flexcpChecked, i, Fg.ColIndex("HaveSerial"), i, Fg.ColIndex("HaveSerial")) = flexChecked Then
                    Fg.Cell(flexcpChecked, i, Fg.ColIndex("Print"), i, Fg.ColIndex("Print")) = flexUnchecked
                Else
                    Fg.Cell(flexcpChecked, i, Fg.ColIndex("Print"), i, Fg.ColIndex("Print")) = flexChecked
                End If

            Next i

    End Select

    If val(Me.TxtQty.text) > 0 Then
        Fg.Cell(flexcpText, Fg.FixedRows, Fg.ColIndex("Qty"), Fg.Rows - 1, Fg.ColIndex("Qty")) = val(Me.TxtQty.text)
    End If

    Me.Lbl(4).Caption = GetCheckedCount
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdPrint_Click()
    Dim Msg As String

    If val(Me.Lbl(4).Caption) > 0 Then
        Load FrmSetting
        FrmSetting.PrintType = 1
        FrmSetting.show vbModal
    Else
        Msg = "ÌÃ»  ÕœÌœ «·√’‰«› «·„—«œ ÿ»«⁄… √ﬂÊ«œÂ«"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Me.Lbl(4).Caption = GetCheckedCount
End Sub

Private Sub Fg_SelChange()
    Dim SngValue As Single
    Dim IntCheckState As VSFlex8UCtl.CellCheckedSettings

    If Fg.Col = Fg.ColIndex("Qty") And Fg.ColSel = Fg.ColIndex("Qty") Then
        If Trim(Fg.TextMatrix(Fg.Row, Fg.Col)) <> "" Then
            SngValue = val(Fg.TextMatrix(Fg.Row, Fg.Col))

            If SngValue > 0 Then
                Fg.Cell(flexcpText, Fg.Row, Fg.Col, Fg.RowSel, Fg.ColSel) = SngValue
            End If
        End If

    ElseIf Fg.Col = Fg.ColIndex("Print") And Fg.ColSel = Fg.ColIndex("Print") Then
        IntCheckState = Fg.Cell(flexcpChecked, Fg.Row, Fg.Col)
        Fg.Cell(flexcpChecked, Fg.Row, Fg.Col, Fg.RowSel, Fg.ColSel) = IntCheckState
    End If

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.Fg

        Select Case .ColKey(Col)

            Case "Print", "Qty"
                Cancel = False

            Case Else
                Cancel = True
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As New ClsBackGroundPic
    Me.Height = 6450
    Me.Width = 8460

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    LoadData

    With Me.CboActions
        .Clear
        .AddItem " ÕœÌœ «·ﬂ·"
        .AddItem "≈·€«¡  ÕœÌœ «·ﬂ·"
        .AddItem " ÕœÌœ «·√’‰«› «· Ï ·Â« ”Ì—Ì«·"
        .AddItem " ÕœÌœ «·√’‰«› «· Ï ·Ì” ·Â« ”Ì—Ì«·"
    End With

    XPCboSearchType.Clear
    XPCboSearchType.AddItem "«”„ «·’‰› "
    XPCboSearchType.AddItem "ﬂÊœ «·’‰› "
    XPCboSearchType.ListIndex = 0
    Resize_Form Me
End Sub

Private Function GetCheckedCount() As Long
    Dim i As Long
    Dim LngChecked As Long

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("Print"), i, .ColIndex("Print")) = flexChecked Then
                LngChecked = LngChecked + 1
            End If

        Next i

    End With

    GetCheckedCount = LngChecked
End Function

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtQty.text, 1)
End Sub

Private Sub XPBtnSearch_Click()
    On Error GoTo ErrTrap
    Dim XGrdNode As VSFlexNode
    Dim LngFindRow As Long
    Dim Msg As String

    If XPCboSearchType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «·Õﬁ· " & Chr(13)
        Msg = Msg + "«·–Ì ”Ì „ «·»ÕÀ ›ÌÂ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPCboSearchType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If XPTxtSearchValue.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «· Ì  —€» ›Ì «·»ÕÀ ⁄‰Â« "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtSearchValue.SetFocus
        Exit Sub
    End If

    If XPChkFullMuch.value = Checked Then
        If XPChkSearchType.value = Checked Then

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemName"), True, True)

                Case 1
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemCode"), True, True)
            End Select

        Else

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemName"), False, True)

                Case 1
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemCode"), False, True)
            End Select

        End If

    Else

        If XPChkSearchType.value = Checked Then

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemName"), True, False)

                Case 1
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemCode"), True, False)
            End Select

        Else

            Select Case XPCboSearchType.ListIndex

                Case 0
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemName"), False, False)

                Case 1
                    LngFindRow = Fg.FindRow(XPTxtSearchValue.text, Fg.FixedRows, Fg.ColIndex("ItemCode"), False, False)
            End Select

        End If
    End If

    If LngFindRow > 0 Then
        '    If XPCboMenuType.ListIndex = 0 Then
        '        If XPOptViewType(0).Value = True Then
        '            Set XGrdNode = Fg.GetNode(LngFindRow)
        '            If Not XGrdNode Is Nothing Then
        '                  XGrdNode.Expanded = True
        '            End If
        '        End If
        '    End If
        Fg.Row = LngFindRow
        Fg.ShowCell LngFindRow, Fg.ColIndex("ItemName")
    Else
        Fg.Row = 1
        Msg = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï Â–« «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    Exit Sub
ErrTrap:
End Sub
