VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddItemAttachs 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈÷«›… «·√’‰«› «·„·Õﬁ…"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "FrmAddItemAttachs.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   7710
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2685
      Left            =   0
      TabIndex        =   1
      Top             =   1650
      Width           =   7695
      _cx             =   13573
      _cy             =   4736
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
      AllowUserResizing=   0
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAddItemAttachs.frx":038A
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
      ExplorerBar     =   0
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
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·’‰›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1485
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   7695
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "1"
         Top             =   1110
         Width           =   1365
      End
      Begin VB.CheckBox ChSelectAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   375
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„Ì…"
         Height          =   225
         Index           =   12
         Left            =   3900
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   9
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   8
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   750
         Width           =   4965
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ã„Ê⁄… «·’‰› :"
         Height          =   225
         Index           =   7
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   6
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√’‰«› «·„Õ·ﬁ… :"
         Height          =   225
         Index           =   5
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   4
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   510
         Width           =   4965
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰› :"
         Height          =   225
         Index           =   2
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   3
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·’‰› :"
         Height          =   225
         Index           =   1
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰› :"
         Height          =   225
         Index           =   0
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   510
         Width           =   1125
      End
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   4620
      Width           =   855
      _ExtentX        =   1508
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
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   4620
      Width           =   885
      _ExtentX        =   1561
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
   Begin VB.Label lblqty 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   225
      Index           =   11
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·√’‰«› «·„Õœœ…:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   10
      Left            =   5910
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -30
      X2              =   7695
      Y1              =   4530
      Y2              =   4545
   End
End
Attribute VB_Name = "FrmAddItemAttachs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_UserCancelled As Boolean

Private Sub Check1_Click()

End Sub

Private Sub ChSelectAll_Click()
  Dim i As Integer

    If ChSelectAll.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
                
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If
  Me.lbl(11).Caption = ModFgLib.GetFgCheckCount(FG, FG.ColIndex("Select"))
  '  Me.lbl(14).Caption = Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
'Me.lbl(14).Caption = val(Calculate_TotalSelected)
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
    UserCancelled = True
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String

    If val(Me.lbl(11).Caption) = 0 Then
        Msg = "ÌÃ» ≈Œ Ì«— ’‰› Ê«Õœ ⁄·Ï «·√ﬁ·...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Me.Hide
    UserCancelled = False
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With Me.FG

        Select Case .ColKey(Col)

            Case "Select"
                Me.lbl(11).Caption = ModFgLib.GetFgCheckCount(FG, FG.ColIndex("Select"))
        End Select

    End With

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.FG

        Select Case FG.ColKey(Col)

'            Case "Select", "AttachItemQty", "AttachItemPrice"
        Case "Select"
                Cancel = False

            Case Else
                Cancel = True
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion

    With Me.FG
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Set CMDOK.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
    Set cmdCancel.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
End Sub

Public Sub LoadItemData(LngItemID As Long, Optional Qty As Double)
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblItems.GroupID, dbo.TblItemsAttach.TableID,dbo.TblItemsAttach.AttachItemID," & "TblItems_1.ItemCode AS AttachItemCode, TblItems_1.ItemName AS AttachItemName," & "TblItems_1.GroupID AS AttachGroupID, dbo.TblItemsAttach.AttachItemQty, " & "dbo.TblItemsAttach.AttachItemPrice "
        StrSQL = StrSQL + " FROM dbo.TblItemsAttach INNER JOIN "
        StrSQL = StrSQL + "dbo.TblItems ON dbo.TblItemsAttach.ItemID = dbo.TblItems.ItemID INNER JOIN "
        StrSQL = StrSQL + "dbo.TblItems TblItems_1 ON dbo.TblItemsAttach.AttachItemID = TblItems_1.ItemID"
        StrSQL = StrSQL + " Where dbo.TblItems.ItemID=" & LngItemID
    Else
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Me.lbl(3).Caption = rs("ItemID").value
        Me.lbl(6).Caption = rs("ItemCode").value
        Me.lbl(4).Caption = rs("ItemName").value
        Me.lbl(9).Caption = rs.RecordCount

        With Me.FG
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("AttachItemID")) = IIf(IsNull(rs("AttachItemID").value), "", rs("AttachItemID").value)
                .TextMatrix(i, .ColIndex("AttachItemCode")) = IIf(IsNull(rs("AttachItemCode").value), "", rs("AttachItemCode").value)
                .TextMatrix(i, .ColIndex("AttachItemName")) = IIf(IsNull(rs("AttachItemName").value), "", rs("AttachItemName").value)
                .TextMatrix(i, .ColIndex("AttachItemQty")) = IIf(IsNull(rs("AttachItemQty").value), 0, rs("AttachItemQty").value) * Qty
                .TextMatrix(i, .ColIndex("OldQty")) = IIf(IsNull(rs("AttachItemQty").value), 0, rs("AttachItemQty").value) * Qty
               If SystemOptions.attacheditemsisfree = True Then
               .TextMatrix(i, .ColIndex("AttachItemPrice")) = 0
               Else
                .TextMatrix(i, .ColIndex("AttachItemPrice")) = IIf(IsNull(rs("AttachItemPrice").value), "", rs("AttachItemPrice").value)
                End If
                '.TextMatrix(I, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Or UnloadMode = VBRUN.QueryUnloadConstants.vbAppTaskManager Then
        Me.Hide
        UserCancelled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Public Property Get UserCancelled() As Boolean
    UserCancelled = m_UserCancelled
End Property

Public Property Let UserCancelled(ByVal vNewValue As Boolean)
    m_UserCancelled = vNewValue
End Property

Private Sub TxtQty_Change()
    Dim i As Long
    
    For i = 1 To FG.Rows - 1
        FG.TextMatrix(i, FG.ColIndex("AttachItemQty")) = val(Txtqty) * val(FG.TextMatrix(i, FG.ColIndex("OldQty")))
    Next

End Sub

