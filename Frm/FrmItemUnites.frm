VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemUnites 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈⁄œ«œ ÊÕœ«  ’‰›"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtRowNumber 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   150
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox TxtUnitPurPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   90
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5190
      Width           =   795
   End
   Begin VB.TextBox TxtUnitSalesPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   930
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5190
      Width           =   795
   End
   Begin VB.CheckBox ChkDef 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÊÕœ… ≈› —«÷Ì…"
      Height          =   315
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox TxtUnitFactor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1740
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5190
      Width           =   1785
   End
   Begin MSDataListLib.DataCombo DcboUnits 
      Height          =   315
      Left            =   3570
      TabIndex        =   3
      Top             =   5190
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid FgUnites 
      Height          =   4125
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   6795
      _cx             =   11986
      _cy             =   7276
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmItemUnites.frx":0000
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
   Begin MSDataListLib.DataCombo DcboItems1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   330
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   510
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5970
      Width           =   5610
      _cx             =   9895
      _cy             =   900
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
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
         Height          =   420
         Index           =   3
         Left            =   780
         TabIndex        =   8
         Top             =   90
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   741
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ›Ÿ"
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
         ButtonImage     =   "FrmItemUnites.frx":01EC
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   420
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   741
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
         ButtonImage     =   "FrmItemUnites.frx":0586
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   90
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14737632
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmItemUnites.frx":0920
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   105
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   9.75
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmItemUnites.frx":0CBA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   0
      Left            =   750
      TabIndex        =   15
      Top             =   5550
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈÷«›…"
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
      ButtonImage     =   "FrmItemUnites.frx":1054
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   1
      Left            =   30
      TabIndex        =   16
      Top             =   5550
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
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
      ButtonImage     =   "FrmItemUnites.frx":13EE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "”⁄— «·‘—«¡"
      Height          =   255
      Index           =   5
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4890
      Width           =   795
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "”⁄— «·»Ì⁄"
      Height          =   255
      Index           =   4
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4890
      Width           =   795
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÊÕœ… ≈› —«÷Ì…"
      Height          =   255
      Index           =   3
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4890
      Width           =   1335
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄·«ﬁ… „⁄ «·ÊÕœ… «·”«»ﬁ…"
      Height          =   255
      Index           =   1
      Left            =   1770
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4890
      Width           =   1755
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·ÊÕœ…"
      Height          =   255
      Index           =   0
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4890
      Width           =   1905
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   2
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   330
      Width           =   855
   End
End
Attribute VB_Name = "FrmItemUnites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSearch(1) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
Select Case Index
    Case 0
        AddNewRow
    Case 1
        RemoveFgRow
    Case 2
        Unload Me
    Case 3
        SaveData_Unites
End Select
End Sub
Private Function Get_SmallUnitFactor(IntBegineRow As Integer) As Double
Dim DblRes As Double
Dim I As Integer

DblRes = 1
With Me.FgUnites
    For I = IntBegineRow + 1 To .Rows - 1 Step 1
        If .TextMatrix(I, .ColIndex("UnitID")) <> "" Then
            DblRes = (DblRes * IIf(Val(.TextMatrix(I, .ColIndex("UnitFactor"))) = _
            0, 1, Val(.TextMatrix(I, .ColIndex("UnitFactor")))))
        Else
            Exit For
        End If
    Next I
End With
Get_SmallUnitFactor = DblRes
End Function
Private Sub DcboItems1_Change()
Dim Rs As ADODB.Recordset
Dim StrSQL As String
Dim I As Integer

If Val(Me.DcboItems1.BoundText) = 0 Then
    Me.FgUnites.Rows = Me.FgUnites.FixedRows
    Exit Sub
End If
StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & _
"TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & _
"TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & _
"TblItemsUnits.UnitID "
StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & Val(Me.DcboItems1.BoundText)
Set Rs = New ADODB.Recordset
Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Not (Rs.BOF Or Rs.EOF) Then
    With Me.FgUnites
        .Rows = Me.FgUnites.FixedRows + Rs.RecordCount
        Rs.MoveFirst
        For I = .FixedRows To .Rows - 1
            If Rs("DefaultUnit").Value = 1 Then
                .Cell(flexcpChecked, I, .ColIndex("DefaultUnit")) = flexChecked
            Else
                 .Cell(flexcpChecked, I, .ColIndex("DefaultUnit")) = flexUnchecked
            End If
            .TextMatrix(I, .ColIndex("UnitID")) = IIf(IsNull(Rs("UnitID").Value), "", Rs("UnitID").Value)
            .TextMatrix(I, .ColIndex("UnitName")) = IIf(IsNull(Rs("UnitName").Value), "", Rs("UnitName").Value)
            .TextMatrix(I, .ColIndex("UnitFactor")) = IIf(IsNull(Rs("UnitFactor").Value), "", Rs("UnitFactor").Value)
            
            
            .TextMatrix(I, .ColIndex("UnitSalesPrice")) = IIf(IsNull(Rs("UnitSalesPrice").Value), "", Rs("UnitSalesPrice").Value)
            .TextMatrix(I, .ColIndex("UnitPurPrice")) = IIf(IsNull(Rs("UnitPurPrice").Value), "", Rs("UnitPurPrice").Value)
            
            .TextMatrix(I, .ColIndex("SecOrder")) = IIf(IsNull(Rs("SecOrder").Value), "", Rs("SecOrder").Value)
            WriteDes CLng(I)
            Rs.MoveNext
        Next I
        .AutoSize 0, .Cols - 1, False
    End With
Else
    Me.FgUnites.Rows = Me.FgUnites.FixedRows
    Exit Sub
End If
Rs.Close
Set Rs = Nothing
End Sub

Private Sub DcboItems1_Click(Area As Integer)
DcboItems1_Change
End Sub

Private Sub FgUnites_DblClick()
With Me.FgUnites
    If .Row <= 0 Then Exit Sub
    If .Col = -1 Then Exit Sub
    
    Me.TxtRowNumber.text = .Row
    If .Cell(flexcpChecked, .Row, .ColIndex("DefaultUnit")) = flexChecked Then
        Me.ChkDef.Value = vbChecked
    Else
        Me.ChkDef.Value = vbUnchecked
    End If
    Me.DcboUnits.BoundText = _
        .TextMatrix(.Row, .ColIndex("UnitID"))
    Me.TxtUnitFactor.text = _
        .TextMatrix(.Row, .ColIndex("UnitFactor"))
    Me.TxtUnitSalesPrice.text = _
        .TextMatrix(.Row, .ColIndex("UnitSalesPrice"))
    Me.TxtUnitPurPrice.text = _
        .TextMatrix(.Row, .ColIndex("UnitPurPrice"))

End With
End Sub


Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
Dim Grdback  As ClsBackGroundPic
Set Grdback = New ClsBackGroundPic
With Me.FgUnites
    .Rows = .FixedRows
    Set .WallPaper = Grdback.Picture
    .AutoSize 0, .Cols - 1, False
    .ExtendLastCol = True
    .ExplorerBar = flexExSortShowAndMove
    .RowHeightMin = 300
End With
Set Dcombos = New ClsDataCombos
Dcombos.GetItemsNames Me.DcboItems1
Set cSearch(0) = New clsDCboSearch
Set cSearch(0).Client = Me.DcboItems1

Dcombos.GetItemsUnits Me.DcboUnits
Set cSearch(1) = New clsDCboSearch
Set cSearch(1).Client = Me.DcboUnits

Resize_Form Me
End Sub

Private Function Get_DefalutUnitFactor(IntBegineRow As Integer, IntDefalutRow As Integer) As Double
'Aim:
'Argument:
'
Dim DblRes As Double
Dim I As Integer
Dim BolCalAsc As Boolean
Dim IntForStep As Integer
If IntBegineRow < IntDefalutRow Then
    BolCalAsc = True
    IntForStep = 1
ElseIf IntBegineRow > IntDefalutRow Then
    BolCalAsc = False
    IntForStep = -1
ElseIf IntBegineRow = IntDefalutRow Then
    Get_DefalutUnitFactor = 1
    Exit Function
End If
DblRes = 1
With Me.FgUnites
    If BolCalAsc = True Then
        For I = IntBegineRow + 1 To IntDefalutRow Step IntForStep
            If .TextMatrix(I, .ColIndex("UnitID")) <> "" Then
                DblRes = (DblRes * IIf(Val(.TextMatrix(I, .ColIndex("UnitFactor"))) = 0, 1, Val(.TextMatrix(I, .ColIndex("UnitFactor")))))
            Else
                Exit For
            End If
        Next I
    Else
        For I = IntBegineRow To IntDefalutRow + 1 Step IntForStep
            If .TextMatrix(I, .ColIndex("UnitID")) <> "" Then
                DblRes = (DblRes * IIf(Val(.TextMatrix(I, .ColIndex("UnitFactor"))) = 0, 1, Val(.TextMatrix(I, .ColIndex("UnitFactor")))))
            Else
                Exit For
            End If
        Next I
    End If
End With
If BolCalAsc = True Then
    Get_DefalutUnitFactor = DblRes
Else
    Get_DefalutUnitFactor = (1 / DblRes)
End If
End Function
Private Sub Form_Unload(Cancel As Integer)
For I = LBound(cSearch) To UBound(cSearch)
    Set cSearch(I) = Nothing
Next I
Erase cSearch
End Sub

Private Sub TxtUnitFactor_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitFactor.text, 0)
End Sub
Private Sub AddNewRow()
Dim Msg As String
Dim LngRow As Long
Dim LngFindRow As Long

If Val(Me.DcboUnits.BoundText) = 0 Then
    Msg = "ÌÃ»  ÕœÌœ «·ÊÕœ…...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
If Val(Me.TxtRowNumber.text) = 0 Then
    LngFindRow = FgUnites.FindRow(Val(Me.DcboUnits.BoundText), _
    FgUnites.FixedRows, FgUnites.ColIndex("UnitID"), False, True)
    If LngFindRow <> -1 Then
        Msg = "·«Ì„ﬂ‰  ﬂ—«— «·ÊÕœ… Ì«⁄„ «·»«‘«...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
End If
If Val(Me.TxtUnitFactor.text) = 0 Then
    Msg = "ÌÃ»  ÕœÌœ ⁄·«ﬁ… «·ÊÕœ… ...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If

If Val(Me.TxtRowNumber.text) <> 0 Then
    LngRow = Val(Me.TxtRowNumber.text)
Else
    Me.FgUnites.Rows = Me.FgUnites.Rows + 1
    LngRow = Me.FgUnites.Rows - 1
End If
If LngRow = 1 Then
    If Val(Me.TxtUnitFactor.text) > 1 Then
        Msg = "›Ï Õ«·… «‰  ﬂÊ‰ Â–Â «Ê· ÊÕœ… ··’‰› ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ „⁄«„· «· ÕÊÌ· «ﬂ»— „‰ Ê«Õœ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.TxtUnitFactor.text = 1
    End If
End If
With Me.FgUnites
    If Me.ChkDef.Value = vbChecked Then
        .Cell(flexcpChecked, .FixedRows, .ColIndex("DefaultUnit"), _
            .Rows - 1, .ColIndex("DefaultUnit")) = flexUnchecked
        .Cell(flexcpChecked, LngRow, .ColIndex("DefaultUnit")) = flexChecked
    End If
    .TextMatrix(LngRow, .ColIndex("UnitID")) = Me.DcboUnits.BoundText
    .TextMatrix(LngRow, .ColIndex("UnitName")) = Me.DcboUnits.text
    .TextMatrix(LngRow, .ColIndex("UnitFactor")) = Format(Val(Me.TxtUnitFactor.text), "0.000")
    .TextMatrix(LngRow, .ColIndex("UnitSalesPrice")) = Val(Me.TxtUnitSalesPrice.text)
    .TextMatrix(LngRow, .ColIndex("UnitPurPrice")) = Val(Me.TxtUnitPurPrice.text)
    .TextMatrix(LngRow, .ColIndex("SecOrder")) = _
        Val(.TextMatrix(LngRow - 1, .ColIndex("SecOrder"))) + 1
     WriteDes LngRow
    .AutoSize 0, .Cols - 1, False
End With

Me.ChkDef.Value = vbUnchecked

Me.DcboUnits.BoundText = ""
Me.TxtUnitFactor.text = ""
Me.TxtUnitSalesPrice.text = ""
Me.TxtUnitPurPrice.text = ""

Me.TxtRowNumber.text = ""
Me.DcboUnits.SetFocus
End Sub

Private Sub SaveData_Unites()

Dim Rs As ADODB.Recordset
Dim StrSQL As String
Dim I As Long
Dim Msg As String
Dim LngCount As Long
Dim IntDefUnitRow As Integer
If Val(Me.DcboItems1.BoundText) = 0 Then
    Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰› ...!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If
LngCount = ItemsInGrid(FgUnites, FgUnites.ColIndex("UnitID"))
If LngCount = 0 Then
    Msg = "ÌÃ» ≈œŒ«· ÊÕœ… ⁄·Ï «·√ﬁ· ....!!!"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
ElseIf Me.FgUnites.FixedRows + 1 = Me.FgUnites.Rows Then
    With Me.FgUnites
       .Cell(flexcpChecked, 1, .ColIndex("DefaultUnit")) = flexChecked
    End With
Else
    If GetFgCheckCount(FgUnites, FgUnites.ColIndex("DefaultUnit")) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ ÊÕœ… ≈› —«÷Ì… ··’‰› ....!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
End If

For I = Me.FgUnites.FixedRows To Me.FgUnites.Rows - 1
    If FgUnites.Cell(flexcpChecked, I, FgUnites.ColIndex("DefaultUnit")) = flexChecked Then
        IntDefUnitRow = I
        Exit For
    End If
Next I
StrSQL = "Delete  From TblItemsUnits Where ItemID=" & Val(Me.DcboItems1.BoundText)
Cn.Execute StrSQL, , adExecuteNoRecords
Set Rs = New ADODB.Recordset
Rs.Open "TblItemsUnits", Cn, adOpenStatic, adLockOptimistic, adCmdTable
With FgUnites
    For I = Me.FgUnites.FixedRows To Me.FgUnites.Rows - 1
        Rs.AddNew
            Rs("ItemID").Value = Val(Me.DcboItems1.BoundText)
            Rs("UnitID").Value = Val(.TextMatrix(I, .ColIndex("UnitID")))
            Rs("UnitFactor").Value = Val(.TextMatrix(I, .ColIndex("UnitFactor")))
            Rs("UnitSalesPrice").Value = Val(.TextMatrix(I, .ColIndex("UnitSalesPrice")))
            Rs("UnitPurPrice").Value = Val(.TextMatrix(I, .ColIndex("UnitPurPrice")))
            If .Cell(flexcpChecked, I, .ColIndex("DefaultUnit")) = flexChecked Then
                Rs("DefaultUnit").Value = 1
            Else
                Rs("DefaultUnit").Value = 0
            End If
            Rs("SecOrder").Value = Val(.TextMatrix(I, .ColIndex("SecOrder")))
            .TextMatrix(I, .ColIndex("FactorByDefaultUnit")) = Format(Get_DefalutUnitFactor(CInt(I), IntDefUnitRow), "0.000")
            Rs("FactorByDefaultUnit").Value = Val(.TextMatrix(I, .ColIndex("FactorByDefaultUnit")))
            
            .TextMatrix(I, .ColIndex("FactorBySmallUnit")) = Format(Get_SmallUnitFactor(CInt(I)), "0.000")
            Rs("FactorBySmallUnit").Value = Val(.TextMatrix(I, .ColIndex("FactorBySmallUnit")))
            
        Rs.update
    Next I
End With
Rs.Close
Set Rs = Nothing
Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub TxtUnitPurPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitPurPrice.text, 0)
End Sub

Private Sub TxtUnitSalesPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtUnitSalesPrice.text, 0)
End Sub

Private Sub RemoveFgRow()
With Me.FgUnites
    If .Row <= 0 Then Exit Sub
    .RemoveItem .Row
End With
End Sub

Private Sub WriteDes(LngRow As Long)
Dim StrTemp1 As String
Dim StrTemp2 As String

With Me.FgUnites
    If LngRow = 1 Then
        .TextMatrix(LngRow, .ColIndex("FactorDes")) = "«·ÊÕœ… «·√Ê·Ï"
    Else
        StrTemp1 = .TextMatrix(LngRow - 1, .ColIndex("UnitName"))
        StrTemp2 = StrTemp1 & "=" & .TextMatrix(LngRow, .ColIndex("UnitFactor")) _
        & .TextMatrix(LngRow, .ColIndex("UnitName"))
        .TextMatrix(LngRow, .ColIndex("FactorDes")) = StrTemp2
    End If
End With
End Sub
