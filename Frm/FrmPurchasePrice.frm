VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPurchasePrice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√”⁄«— «·‘—«¡"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "FrmPurchasePrice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2610
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   6240
      _cx             =   11007
      _cy             =   4604
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmPurchasePrice.frx":038A
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
      Index           =   6
      Left            =   90
      TabIndex        =   0
      Top             =   4290
      Width           =   825
      _ExtentX        =   1455
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
      ButtonImage     =   "FrmPurchasePrice.frx":043F
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label XPLblLargePrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4020
      Width           =   915
   End
   Begin VB.Label XPLblLessPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4380
      Width           =   915
   End
   Begin VB.Label XPLblCusName1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4020
      Width           =   2385
   End
   Begin VB.Label XPLblItemID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   690
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label XPLblSerial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2130
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   990
      Width           =   4005
   End
   Begin VB.Label XPLblItemName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   255
      Index           =   1
      Left            =   5370
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   690
      Width           =   885
   End
   Begin VB.Label XPLblCusName2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4380
      Width           =   2385
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «”„ «·„Ê—œ :"
      Height          =   255
      Index           =   5
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4410
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " √ﬁ· ”⁄— :"
      Height          =   255
      Index           =   8
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4410
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «”„ «·„Ê—œ :"
      Height          =   255
      Index           =   9
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4050
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " √⁄·Ï ”⁄— : "
      Height          =   255
      Index           =   10
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4050
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " „ﬁ«—‰… «·√”⁄«— "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   645
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "FrmPurchasePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset

Private Sub Cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap
    Dim RecordNum As Integer
    Dim ChildNum As Integer
    Dim StrSQL As String
    Dim LngRow As Long
    Set rs = New ADODB.Recordset
    Set RsTemp = New ADODB.Recordset
    StrSQL = "select * From TblItems where ItemID=" & XPLblItemID.Caption
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            XPLblSerial.ForeColor = &HFF&
            XPLblSerial.Caption = "Â–« «·’‰› Ì ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
        Else
            XPLblSerial.ForeColor = &HFF0000
            XPLblSerial.Caption = "Â–« «·’‰› ·« Ì ⁄«„· »‰Ÿ«„ «·”Ì—Ì«·"
        End If
    End If

    RsTemp.Close
    StrSQL = "SELECT CusJuncItem.ID, CusJuncItem.CusID, CusJuncItem.ItemID, CusJuncItem.ItemPrice, " & "TblCustemers.CusName FROM TblCustemers INNER JOIN CusJuncItem ON TblCustemers.CusID=CusJuncItem.CusID"
    StrSQL = StrSQL + " WHERE CusJuncItem.ItemID=" & XPLblItemID.Caption
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.FG
        .Rows = .FixedRows
        .GridLines = flexGridNone
        .Redraw = False
        .OutlineCol = 0

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst

            For RecordNum = 1 To rs.RecordCount
                .Rows = .Rows + 1
                .Cell(flexcpText, .Rows - 1, .ColIndex("Custemer")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .Cell(flexcpText, .Rows - 1, .ColIndex("Price")) = IIf(IsNull(rs("ItemPrice").value), "", rs("ItemPrice").value)
                .NodeClosedPicture = mdifrmmain.ImgLstTree.ListImages("Close").Picture
                .NodeOpenPicture = mdifrmmain.ImgLstTree.ListImages("Root").Picture
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
                .MergeRow(.Rows - 1) = True
                StrSQL = "select * From JuncPrice where juncID=" & rs("ID").value
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                For ChildNum = 1 To RsTemp.RecordCount
                    .AddItem "" & vbTab & "" & vbTab & RsTemp("From").value & vbTab & RsTemp("To").value & vbTab & RsTemp("Price").value
                    .Cell(flexcpBackColor, .Rows - 1, .ColIndex("Form"), .Rows - 1, .ColIndex("Cost")) = &HE2E9E9
                    .RowOutlineLevel(.Rows - 1) = 2
                    RsTemp.MoveNext
                Next ChildNum

                RsTemp.Close
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            .OutlineBar = flexOutlineBarCompleteLeaf
            .Ellipsis = flexEllipsisEnd
            .Outline 1
            .ColWidth(.ColIndex("TreeColomn")) = 270
            rs.MoveFirst
        Else
            .Rows = .FixedRows + 1
        End If

    End With

    FG.Redraw = flexRDDirect
    FG.AutoSize 0, FG.Cols - 1, False
    XPLblLargePrice.Caption = FG.Aggregate(flexSTMax, FG.FixedRows, FG.ColIndex("Price"), FG.Rows - 1, FG.ColIndex("Price"))
    LngRow = FG.FindRow(XPLblLargePrice.Caption, FG.FixedRows, FG.ColIndex("Price"))
    XPLblCusName1.Caption = FG.TextMatrix(LngRow, FG.ColIndex("Custemer"))
    XPLblLessPrice.Caption = FG.Aggregate(flexSTMin, FG.FixedRows, FG.ColIndex("Price"), FG.Rows - 1, FG.ColIndex("Price"))
    LngRow = FG.FindRow(XPLblLessPrice.Caption, FG.FixedRows, FG.ColIndex("Price"))
    XPLblCusName2.Caption = FG.TextMatrix(LngRow, FG.ColIndex("Custemer"))
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim GrdBak As New ClsBackGroundPic

    FormPostion Me, GetPostion

    With FG
        .Cell(flexcpPicture, 0, .ColIndex("Custemer")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Price")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Form")) = mdifrmmain.ImgLstTree.ListImages("From").Picture
        .Cell(flexcpPicture, 0, .ColIndex("To")) = mdifrmmain.ImgLstTree.ListImages("To").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Cost")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Set FG.WallPaper = GrdBak.Picture
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    If RsTemp.State = adStateOpen Then
        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If RsTemp.EditMode <> adEditNone Then
                RsTemp.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set RsTemp = Nothing
    Exit Sub
ErrTrap:
End Sub

