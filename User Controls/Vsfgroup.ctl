VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl vsfGroup 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   KeyPreview      =   -1  'True
   PropertyPages   =   "Vsfgroup.ctx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   6015
   ToolboxBitmap   =   "Vsfgroup.ctx":0012
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   2670
      Top             =   180
   End
   Begin MSComctlLib.ImageList ImgListTreeIcons 
      Left            =   4500
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vsfgroup.ctx":010C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vsfgroup.ctx":04A6
            Key             =   "Selected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vsfgroup.ctx":0840
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Vsfgroup.ctx":0BDA
            Key             =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TrvGroups 
      Height          =   3765
      Left            =   3450
      TabIndex        =   2
      Top             =   660
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   6641
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   570
      ScaleHeight     =   360
      ScaleWidth      =   1260
      TabIndex        =   0
      Tag             =   "Hello"
      Top             =   270
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2835
      Left            =   150
      TabIndex        =   1
      Top             =   1080
      Width           =   2640
      _cx             =   4657
      _cy             =   5001
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
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
      Begin VB.Shape ShpPointer 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   465
         Left            =   540
         Top             =   1290
         Width           =   1695
      End
      Begin VB.Image ImgPointer 
         Height          =   525
         Left            =   1110
         Picture         =   "Vsfgroup.ctx":0F74
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.Label picDrag 
      BackColor       =   &H00808080&
      Height          =   3555
      Left            =   3150
      TabIndex        =   3
      Top             =   810
      Width           =   60
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "MnuOptions"
      Visible         =   0   'False
      Begin VB.Menu MnuPutBookMark 
         Caption         =   "Put BookMark in this Group"
      End
   End
End
Attribute VB_Name = "vsfGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'######################################################
'„‰ ≈‰ «Ã ‘—ﬂ… »«Ì  ··»—„ÃÌ« 
'  «·„»—„ÃÌ‰ «·–Ï ⁄„·Ê« ›Ï  ⁄œÌ· «·√œ«…
'Ê ÿÊÌ—Â«
'«Ì„‰ „Õ„œ ⁄„«—…
'«Ì„«‰ ⁄»œ «·—«“ﬁ „Õ„œ
'—»«» Õ”‰ «»—«ÂÌ„
'######################################################
'--------------------------------------------------------
'Events Declarations
Public Event GridDblClick(Row As Long, Col As Long)

'--------------------------------------------------------
Public Enum PointersTypes
    ArrowPointer
    ShapeBorderPointer
End Enum

' API declarations
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const DFC_BUTTON = 4

Private Const DFCS_BUTTONPUSH = &H10

Private Declare Function DrawFrameControl _
                Lib "user32" (ByVal hDC As Long, _
                              lpRect As RECT, _
                              ByVal un1 As Long, _
                              ByVal un2 As Long) As Long

Private Declare Function SetCapture _
                Lib "user32" (ByVal hWnd As Long) As Long

'--------------------------------------------------------
' private declarations
Private Type POINTSGL
    X As Single
    Y As Single
End Type

Private Type GROUPINFO
    ctl As PictureBox
    Text As String
    GroupDataType As VSFlex8UCtl.DataTypeSettings
End Type

Private Type ColFilter
    ColKey As String
    ColDataType As String
    FilterValue As String
End Type

Private Const CLR_BTNFACE = &H8000000F

Private Const CLR_BTNSHADOW = &H80000010

Private Const CLR_BTNHILITE = &H80000014

Private HELPMSG As String

Private Const DRAG_TOLERANCE = 100 ' Twips

'--------------------------------------------------------
Private LngPointerInterval As Long

' variables
' mouse control
Private m_bCapture As Boolean   ' mouse captured?

Private m_bDragging As Boolean  ' dragging control?

Private m_ptDown As POINTSGL    ' where was the click

Private m_ptControl As POINTSGL ' original coordinates

Private m_iGroups As Integer    ' how many groups do we have

Private m_GroupInfo() As GROUPINFO ' group information vector

Private m_SetRTL  As Boolean 'iF The Control Will be Right To Left Or NOT

Private m_LngTotalOnColKey As String 'the index of the col Which we will Sum its Values

Private m_SQL As String 'The Main SQL Statmeent which load data to the control

Private m_ShowTreeGroups As Boolean 'Show Or Hide the treeView Control

Private m_SeparatorColor As Single

Private m_PointerType As PointersTypes

Private Function FindColumn(s$) As Integer
    ' locate column based on header text
    Dim i%

    For i = 0 To FG.Cols - 1

        If FG.Cell(flexcpTextDisplay, 0, i) = s Then
            FindColumn = i
            Exit Function
        End If

    Next
    
    ' this should never happen
    FindColumn = -1

End Function

Private Sub UpdateGrid()
    Dim j As Long
    Dim lngCount As Long
    Dim SngForeColor As Single
    Dim XNode As VSFlex8UCtl.VSFlexNode
    '-------------------------------------
    SngForeColor = &H80&
    '-------------------------------------
    ' redraw is off to speed things up
    FG.Redraw = True
    FG.SubtotalPosition = flexSTAbove
    ' move groups to left
    Dim i As Long, Col As Long
    FG.Subtotal flexSTClear

    For i = 0 To m_iGroups - 1
        Col = FindColumn(m_GroupInfo(i).Text)
        FG.ColPosition(Col) = i
    Next

    'hide groups, make sure they're all sortable
    For i = 0 To m_iGroups - 1
        FG.ColHidden(i) = True

        If FG.ColSort(i) = flexSortNone Then
            FG.ColSort(i) = flexSortGenericAscending
        End If

    Next
    
    'show non-groups
    For i = m_iGroups To FG.Cols - 1
        FG.ColHidden(i) = False
    Next
    
    ' sort
    If FG.Row <> -1 Then
        FG.Select FG.Row, 0, FG.Row, FG.Cols - 1
    End If
    
    FG.Sort = flexSortUseColSort
    ' create groups
    FG.Subtotal flexSTClear
    FG.Redraw = True

    If m_iGroups > 0 Then

        For i = 0 To m_iGroups - 1

            If Me.TotalOnColKey = "" Then
                FG.Subtotal flexSTNone, i, 0, , CLR_BTNFACE, SngForeColor, True, , , True
            Else
                FG.Subtotal flexSTSum, i, FG.ColIndex(Me.TotalOnColKey), "#,###.##", CLR_BTNFACE, GetLevelForeColor(CInt(i)), True, " ≈Ã„«·Ï %s ", , True
            End If

        Next i

        FG.Redraw = True
        ' group them
        FG.Outline m_iGroups - 1
        FG.OutlineCol = m_iGroups
        FG.AutoSize m_iGroups
    End If
    
    ' move text to visible rows
    FG.MousePointer = flexArrowHourGlass

    If m_iGroups > 0 Then

        For i = 1 To FG.Rows - 1

            If FG.IsSubtotal(i) Then
                Dim s$
                Set XNode = FG.GetNode(i)
                s = FG.Cell(flexcpTextDisplay, i, 0)
                FG.Cell(flexcpText, i, 0) = ""
                's = s & " - " & XNode.Children
                FG.Cell(flexcpText, i, m_iGroups) = s & "-" & GetNodeChildCount(XNode)
            End If

        Next

        If Me.ShowTreeGroups = True Then
            LoadTreeGroups
        End If

    Else
        TrvGroups.Nodes.Clear
    End If
    
    FG.MousePointer = flexDefault
    FG.MergeCells = flexMergeSpill

    ' redraw is back on
    FG.Redraw = True
    
End Sub

Private Sub Fg_AfterCollapse(ByVal Row As Long, _
                             ByVal State As Integer)
    PutPointer
End Sub

Private Sub Fg_AfterScroll(ByVal OldTopRow As Long, _
                           ByVal OldLeftCol As Long, _
                           ByVal NewTopRow As Long, _
                           ByVal NewLeftCol As Long)
    PutPointer
End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single, _
                               Cancel As Boolean)
    
    ' if we clicked on a column, start dragging it
    If Button = 1 And Shift = 0 And FG.MouseRow = 0 Then
    
        ' make sure we don't group on everything
        If m_iGroups >= FG.Cols - 1 Then
            Exit Sub
        End If
    
        ' which column are we grouping on?
        Dim Col As Long
        Col = FG.MouseCol
    
        ' confirm that this is a groupable column
        Dim i As Long

        For i = 0 To m_iGroups - 1

            If m_GroupInfo(i).Text = FG.Cell(flexcpTextDisplay, 0, Col) Then
                Cancel = True
                Beep
                Exit Sub
            End If

        Next

        ' UNDONE
        UserControl.MousePointer = vbArrowHourglass
        FG.MousePointer = flexArrowHourGlass
        ' create entry in global array
        i = m_iGroups
        m_iGroups = m_iGroups + 1
        ReDim Preserve m_GroupInfo(i)
    
        'create new group control
        '-------------------------
        Static newCtl As Integer
        newCtl = newCtl + 1
        Load picGroup(newCtl)
        Set m_GroupInfo(i).ctl = picGroup(newCtl)
        m_GroupInfo(i).Text = FG.Cell(flexcpTextDisplay, 0, Col)
        m_GroupInfo(i).GroupDataType = FG.ColDataType(Col)
    
        If FG.ColDataType(Col) = flexDTDate Then
            FG.ColFormat(Col) = "MM/yyyy"
        End If
    
        ' init group control
        With picGroup(newCtl)
            .Tag = i
            .Width = .TextWidth(m_GroupInfo(i).Text) + (2 * FG.RowHeight(0))
            .Height = FG.RowHeight(0) * 1.1
            .Move FG.ColPos(Col), FG.top
            .Font = FG.Font
            .RightToLeft = picGroup(0).RightToLeft
            .ZOrder
        End With
    
        ' save original position (none in this case)
        m_ptControl.X = -1
        m_ptControl.Y = -1
    
        ' start dragging
        m_bCapture = True
        m_bDragging = True
        m_ptDown.X = X - picGroup(newCtl).left
        m_ptDown.Y = FG.top + Y - picGroup(newCtl).top
        picGroup_Paint newCtl
    
        ' this is really cool:
        ' flex got the mouse down, but we want the group control to handle it
        ' so we set Cancel to true and transfer the mouse to the group control
        ' using the SetCapture API.
        Cancel = True

        With picGroup(newCtl)
            .Visible = True
            .SetFocus
            SetCapture .hWnd
        End With

        UserControl.MousePointer = vbDefault
        FG.MousePointer = flexDefault
    End If

End Sub

Private Sub FG_BeforeScroll(ByVal OldTopRow As Long, _
                            ByVal OldLeftCol As Long, _
                            ByVal NewTopRow As Long, _
                            ByVal NewLeftCol As Long, _
                            Cancel As Boolean)
    PutPointer
End Sub

Private Sub Fg_DblClick()
    Dim XGrdNode As VSFlex8UCtl.VSFlexNode
    RaiseEvent GridDblClick(FG.Row, FG.Col)

    With FG

        If .IsSubtotal(.Row) = True Then
            Set XGrdNode = .GetNode(.Row)

            If Not (XGrdNode Is Nothing) Then
                XGrdNode.Expanded = Not XGrdNode.Expanded
            End If
        End If

    End With

End Sub

Private Sub Fg_DragDrop(Source As Control, _
                        X As Single, _
                        Y As Single)

    If Me.SetRTL = True Then
        FG.left = UserControl.ScaleLeft
        FG.Width = X
    ElseIf Me.SetRTL = False Then
        TrvGroups.left = UserControl.ScaleLeft
        TrvGroups.Width = FG.left + X
    End If

    DragDropPostion
End Sub

Private Sub Fg_RowColChange()
    PutPointer
End Sub

Private Sub Fg_SelChange()
    PutPointer
End Sub

Private Sub ImgPointer_DblClick()
    ImgPointer.Visible = False
End Sub

Private Sub picGroup_Click(Index As Integer)

    ' unless we were dragging, revert sort direction
    If (Not m_bDragging) And (m_ptControl.X > -1) Then
        
        ' revert sort direction
        Dim i%
        i = picGroup(Index).Tag

        If FG.ColSort(i) = flexSortGenericDescending Then
            FG.ColSort(i) = flexSortGenericAscending
        Else
            FG.ColSort(i) = flexSortGenericDescending
        End If
        
        ' show the change
        UpdateLayout True
        
    End If

End Sub

Private Sub picGroup_KeyPress(Index As Integer, _
                              KeyAscii As Integer)
    
    ' escape cancels dragging/clicking
    If (KeyAscii = 27) And (m_bCapture = True) Then
        
        ' move control back to its original position
        If m_bDragging Then
        
            ' if the group was still being created (not just dragged), delete it
            If m_ptControl.X < 0 And m_ptControl.Y < 0 Then
                DeleteGroup Index
            
                ' otherwise, move it back to where it was
            Else
                picGroup(Index).Move m_ptControl.X, m_ptControl.Y
            End If
        End If
        
        ' reset state variables
        m_bCapture = False
        m_bDragging = True
    End If

End Sub

Private Sub DeleteGroup(Index As Integer)
    
    ' remove control From the list
    Dim i%, j%
    i = picGroup(Index).Tag

    For j = i To m_iGroups - 2
        m_GroupInfo(j) = m_GroupInfo(j + 1)
    Next

    m_iGroups = m_iGroups - 1
    
    If m_iGroups = 0 Then FG.Outline 1

    ' hide/unload the control
    picGroup(Index).Visible = False

    If Index > 0 Then Unload picGroup(Index)
    
End Sub

Private Sub picGroup_MouseDown(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    ' left button starts dragging
    If Button = 1 Then
    
        ' save dragging information
        m_bCapture = True
        m_bDragging = False
        m_ptDown.X = X
        m_ptDown.Y = Y
        
        ' bring control to top, save its original position
        picGroup(Index).ZOrder
        m_ptControl.X = picGroup(Index).left
        m_ptControl.Y = picGroup(Index).top
    End If

End Sub

Private Sub picGroup_MouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    ' drag control around
    If m_bCapture Then

        With picGroup(Index)
                    
            ' if we are not dragging yet, maybe it's time to start
            If Not m_bDragging Then
                If Abs(X - m_ptDown.X) > DRAG_TOLERANCE Then m_bDragging = True
                If Abs(Y - m_ptDown.Y) > DRAG_TOLERANCE Then m_bDragging = True
            End If
        
            ' if we're dragging, then do it
            If m_bDragging Then
        
                ' get new coordinates
                X = .left + (X - m_ptDown.X)
                Y = .top + (Y - m_ptDown.Y)
            
                ' restrict boundaries
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                If X > UserControl.ScaleWidth - .Width Then X = UserControl.ScaleWidth - .Width
                If Y > UserControl.ScaleHeight - .Height Then Y = UserControl.ScaleHeight - .Height
                If Y > FG.top Then Y = FG.top
        
                ' move the control
                .Move X, Y
            
                ' show where we'd go if we dropped now
                ' UNDONE
            
            End If

        End With

    End If

    picGroup(Index).ToolTipText = "»«·÷€ÿ Â‰« Ì „  — Ì» «·»Ì«‰«  ›Ï «·ÃœÊ·"
End Sub

Private Sub picGroup_MouseUp(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Dim LngMouseCol As Long
    On Error GoTo ErrTrap

    LngMouseCol = FG.MouseCol

    ' if we were dragging,
    ' we may have just moved the group to a new position, or
    ' we may have dropped it back into the grid
    'If LngMouseCol <= -1 Then Exit Sub
    If m_bDragging Then
        
        FG.Redraw = False
        
        ' back into grid, different position
        Y = picGroup(Index).top + Y

        If Y > FG.top Then
            
            ' see which column it was and where the mouse is
            Dim Col%, i%
            Col = FindColumn(m_GroupInfo(picGroup(Index).Tag).Text)
            i = LngMouseCol
            
            'different? move column
            If i <> Col Then
                FG.ColPosition(Col) = i
                'same? switch sort order
            Else

                If FG.ColSort(i) = flexSortGenericAscending Then
                    FG.ColSort(i) = flexSortGenericDescending
                Else
                    FG.ColSort(i) = flexSortGenericAscending
                End If
            End If
            
            ' remove our brand-new group
            DeleteGroup Index
        
        End If
        
        ' either way, show changes
        UpdateLayout True
        
        FG.Redraw = True
    End If

    ' cancel capture no matter what
    m_bCapture = False
ErrTrap:
End Sub

Private Sub picGroup_Paint(Index As Integer)
    
    Dim rc As RECT
    
    With picGroup(Index)
        
        ' draw frame
        rc.top = 0
        rc.left = 0
        rc.right = .Width / Screen.TwipsPerPixelX
        rc.bottom = .Height / Screen.TwipsPerPixelY
        DrawFrameControl .hDC, rc, DFC_BUTTON, DFCS_BUTTONPUSH
        
        ' draw text
        '###############################################################
        'Update By Nour So Support Right TO Left Reading (For Arabic)
        If picGroup(Index).RightToLeft = True Then
            .CurrentX = .ScaleWidth - .TextWidth(Trim$(m_GroupInfo(.Tag).Text))
        Else
            .CurrentX = .ScaleLeft
        End If

        .CurrentY = (.Height - .TextHeight(" ")) / 2.5
        picGroup(Index).Print m_GroupInfo(.Tag).Text

        '##############################################################
        ' draw sort arrow if this is a group already
        If FG.ColWidth(.Tag) = 0 Then
            Dim X As Single, Y As Single, sz As Single
            sz = .Height * (1 / 3)
            X = .Width - sz
            
            ' pointing up
            If FG.ColSort(.Tag) = flexSortGenericDescending Then
                Y = (.Height - sz) / 2 + sz
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNHILITE
                picGroup(Index).Line -(X - sz / 2, Y - sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            
                ' pointing down
            Else
                Y = (.Height - sz) / 2
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNSHADOW
                picGroup(Index).Line -(X - sz / 2, Y + sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            End If
        End If

    End With

End Sub

Private Sub Timer1_Timer()

    LngPointerInterval = LngPointerInterval + 1

    If LngPointerInterval >= 10 Then
        LngPointerInterval = 0
        RemovePointer
    Else

        If Me.PointerType = ArrowPointer Then
            ImgPointer.Visible = Not ImgPointer.Visible
        ElseIf Me.PointerType = ShapeBorderPointer Then
            ShpPointer.Visible = Not ShpPointer.Visible
        End If
    End If

End Sub

Private Sub TrvGroups_DragDrop(Source As Control, _
                               X As Single, _
                               Y As Single)

    If Me.SetRTL = True Then
        FG.left = UserControl.ScaleLeft
        FG.Width = TrvGroups.left + (TrvGroups.Width - X)
    Else
        TrvGroups.left = UserControl.ScaleLeft
        TrvGroups.Width = X
    End If

    DragDropPostion
End Sub

Private Sub TrvGroups_MouseUp(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    Dim XNode As MSComctlLib.Node

    If Button = vbRightButton Then
        Set XNode = TrvGroups.HitTest(X, Y)

        If Not XNode Is Nothing Then
            UserControl.PopupMenu MnuOptions
        End If
    End If

End Sub

Private Sub TrvGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim StrKey As String
    Dim LngRow As Long
    Dim FgNode As VSFlex8UCtl.VSFlexNode
    Dim LngFirstVisibleCol As Long

    If Not Node Is Nothing Then
        StrKey = Node.Key

        If StrKey = "root" Then
            Exit Sub
        Else
            LngRow = val(StrKey)

            If LngRow <> 0 Then
                If FG.IsSubtotal(LngRow) = True Then
                    Set FgNode = FG.GetNode(LngRow)
                    FgNode.Expanded = True
                    FgNode.EnsureVisible
                    LngFirstVisibleCol = GetFirstVisibleCol
                    FG.ShowCell LngRow, LngFirstVisibleCol

                    '-Set the Pointer
                    If Me.PointerType = ArrowPointer Then
                        ImgPointer.Visible = True
                        ImgPointer.Tag = LngRow
                        ShpPointer.Visible = False
                        ShpPointer.Tag = ""
                    Else
                        ShpPointer.Visible = True
                        ShpPointer.Tag = LngRow
                        ImgPointer.Visible = False
                        ImgPointer.Tag = ""
                    End If

                    PutPointer
                End If
            End If
        End If
    End If

End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    'initialize embedded FlexGrid
    FG.SelectionMode = flexSelectionFree
    FG.AllowUserResizing = flexResizeColumns
    FG.OutlineBar = flexOutlineBarComplete
    FG.ExplorerBar = flexExSortAndMove
    FG.ExtendLastCol = True
    picGroup(0).RightToLeft = UserControl.RightToLeft

    'initialize group control based on grid data
    With picGroup(0)
        .Font = FG.Font
        .Height = FG.RowHeight(0)
        .Tag = 0
    End With

    'Set The Tree View Control
    With TrvGroups
        .Appearance = cc3D
        .Indentation = 10
        .BorderStyle = ccNone
        .Checkboxes = False
        .LabelEdit = tvwManual
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        Set .ImageList = ImgListTreeIcons
    End With

    Me.PointerType = ShapeBorderPointer
    picDrag.MousePointer = 9
    picDrag.backcolor = UserControl.backcolor
    picDrag.ZOrder 0
    picDrag.Height = UserControl.ScaleHeight
    ImgPointer.Visible = False
    Me.ShowTreeGroups = True
    UpdateLayout False
End Sub

Private Sub UpdateLayout(dogrid As Boolean)
    Dim SngTemp  As Single
    Dim swap As GROUPINFO
    Dim i%, cnt%, done%
    Dim X As Single, Y As Single, rh As Single
    Dim offsety As Single
    
    ' see how many groups are visible
    cnt = m_iGroups
    
    ' dimension and clear grouping area
    rh = FG.RowHeight(0)
    offsety = rh / 2
    Y = 2 * FG.RowHeight(0)

    If cnt > 1 Then Y = Y + (cnt - 1) * offsety
    Y = UserControl.ScaleHeight - Y

    If Y < 0 Then Y = 0
    FG.Height = Y
    UserControl.Cls
    
    ' if no groups, show helpful message
    If cnt = 0 Then
        UserControl.CurrentX = rh / 2
        UserControl.CurrentY = rh / 2
        UserControl.Print HELPMSG
    End If
    
    ' sort group vector by position (left-to-right)
    While Not done

        done = True

        For i = 0 To cnt - 2

            If m_GroupInfo(i).ctl.left > m_GroupInfo(i + 1).ctl.left Then

                done = False
                swap = m_GroupInfo(i)
                m_GroupInfo(i) = m_GroupInfo(i + 1)
                m_GroupInfo(i + 1) = swap
            End If

        Next

    Wend
    
    ' each control gets and index into the vector
    For i = 0 To cnt - 1
        m_GroupInfo(i).ctl.Tag = i
    Next
    
    ' position group controls
    Y = rh / 2
    X = Y

    For i = 0 To cnt - 1

        With m_GroupInfo(i).ctl
        
            ' move the control
            .Move X, Y
            Y = Y + offsety
            X = X + .Width + rh / 3
        
            ' draw connector
            If i < cnt - 1 Then
                UserControl.Line (X, Y + 2 / 3 * rh)-(X - rh * 2 / 3, Y + 2 / 3 * rh), 0
                UserControl.Line -(X - rh * 2 / 3, Y + rh / 2 - Screen.TwipsPerPixelY), 0
            End If

            ' draw placeholder
            UserControl.Line (.left, .top)-(.left + .Width - Screen.TwipsPerPixelX, .top + .Height - Screen.TwipsPerPixelY), 0, B
        
        End With

    Next
    
    'redraw all controls at their new positions
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next

    UserControl.Refresh
    
    ' update the grid
    '----------------------------------------------
    RemovePointer

    '----------------------------------------------
    If dogrid Then UpdateGrid
     
    If Me.ShowTreeGroups = True Then
        If Me.SetRTL = True Then
            SngTemp = UserControl.ScaleWidth / 4
            '-----------------------------------
            FG.top = (UserControl.ScaleHeight - (FG.Height + 50))
            '-----------------------------------
            FG.left = UserControl.ScaleLeft
            TrvGroups.Width = SngTemp
            FG.Width = UserControl.ScaleWidth - (picDrag.Width + TrvGroups.Width)
            picDrag.left = (FG.left + FG.Width)
            picDrag.top = FG.top
            picDrag.Height = FG.Height

            With TrvGroups
                .top = FG.top
                .Height = FG.Height
                .left = (picDrag.left + picDrag.Width)
                .Width = UserControl.ScaleWidth - (FG.Width + picDrag.Width)
            End With

        ElseIf Me.SetRTL = False Then
            SngTemp = UserControl.ScaleWidth / 4
            '-------------------------------------
            FG.top = (UserControl.ScaleHeight - (FG.Height + 50))
            '-------------------------------------
            TrvGroups.left = UserControl.ScaleLeft
            TrvGroups.Width = SngTemp
            TrvGroups.top = FG.top
            TrvGroups.Height = FG.Height
            picDrag.left = (TrvGroups.left + TrvGroups.Width)
            picDrag.top = FG.top
            picDrag.Height = FG.Height
            FG.left = (picDrag.left + picDrag.Width)
            FG.Width = UserControl.ScaleWidth - (TrvGroups.Width + picDrag.Width)
        End If

    Else
        FG.top = (UserControl.ScaleHeight - (FG.Height + 50))
        FG.left = UserControl.ScaleLeft
        FG.Width = UserControl.ScaleWidth
        
    End If

    'redraw all controls at their new positions (to show sort direction)
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next

    'FG.AutoSize 0, FG.Cols - 1, False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    picDrag.backcolor = PropBag.ReadProperty("SeparatorColor", vbRed)
    Me.SetRTL = PropBag.ReadProperty("SetRTL", True)
End Sub

Private Sub UserControl_Resize()
    UpdateLayout False
End Sub

Public Property Get VSFlexGrid() As VSFlexGrid
    Set VSFlexGrid = FG
End Property

Public Sub update()
    UpdateLayout True
End Sub

Public Property Get SetRTL() As Boolean
Attribute SetRTL.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute SetRTL.VB_UserMemId = -611
    SetRTL = m_SetRTL
End Property

Public Property Let SetRTL(ByVal vNewValue As Boolean)
    Dim i As Integer
    '›Ï Â–Â «·œ«·… ‰ﬁÊ„ »≈⁄œ«œ «·Ã—œ
    'Õ Ï Ì„ﬂ‰‰« ≈” Œœ«„ ›Ï «·⁄—÷ »«··€… «·⁄—»Ì…
    '«Ê «·√‰Ã·“Ì…

    m_SetRTL = vNewValue
    PropertyChanged "SetRTL"

    If m_SetRTL = True Then
        HELPMSG = "ﬁ„ »”Õ» «Ï ⁄„Êœ Â‰« Õ Ï Ì „ «· Ã„Ì⁄ ⁄·Ï «”«” Â–« «·⁄„Êœ"
        UserControl.RightToLeft = True
        picGroup(0).RightToLeft = True

        With FG
            .RightToLeft = True

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        End With

        'Update the TreeView Control to Right TO Left
        Make_RightToLeft TrvGroups
    Else
        HELPMSG = " Drag a column header here to group by that column. "
        UserControl.RightToLeft = False
        picGroup(0).RightToLeft = False

        With FG
            .RightToLeft = False

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End With

        Make_RightToLeft TrvGroups, True
    End If

    UpdateLayout False
End Property

Private Function GetNodeChildCount(XNode As VSFlex8UCtl.VSFlexNode) As Long
    Dim Y  As VSFlexNode, Z As VSFlexNode
    Dim IntCount As Long
    Dim i As Long

    Set Y = XNode.GetNode(flexNTNextSibling)
    Set Z = XNode.GetNode(flexNTLastSibling)

    '----------------------------------------------------------------
    If Not Y Is Nothing Then

        'GetNodeChildCount = Y.Row - (Xnode.Row + 1 + Xnode.Children)
        For i = XNode.Row To Y.Row

            If FG.IsSubtotal(i) = False Then
                IntCount = IntCount + 1
            End If

        Next i

        GetNodeChildCount = IntCount
    Else

        For i = XNode.Row To FG.Rows - 1

            If FG.IsSubtotal(i) = False Then
                IntCount = IntCount + 1
            Else

                If FG.RowOutlineLevel(i) < FG.RowOutlineLevel(XNode.Row) Then
                    Exit For
                End If
            End If

        Next i

        GetNodeChildCount = IntCount
    End If

End Function

Public Function RefreshData() As Boolean
    Dim rs As ADODB.Recordset
    On Error GoTo hErr

    If Not FG.DataSource Is Nothing Then
        Set rs = New ADODB.Recordset
        Set rs = FG.DataSource
        rs.Requery
        FG.Clear flexClearScrollable, flexClearEverything
        FG.Rows = FG.FixedRows
        Set FG.DataSource = rs

        UpdateLayout True
        FG.Redraw = flexRDDirect
        FG.Redraw = True
        m_bDragging = False
        m_ptControl.X = 0
        picGroup_Click 0
        FG.Outline 0
    
    End If

    RefreshData = True
    Exit Function
hErr:
    RefreshData = False
End Function

Public Property Get TotalOnColKey() As String
Attribute TotalOnColKey.VB_ProcData.VB_Invoke_Property = "VsGroupPage"
    TotalOnColKey = m_LngTotalOnColKey
End Property

Public Property Let TotalOnColKey(ByVal vNewValue As String)
    m_LngTotalOnColKey = vNewValue
End Property

Private Function GetLevelForeColor(IntLevel As Integer) As Single
    Dim SngTemp As Single

    If IntLevel >= 1 And IntLevel <= 3 Then
        SngTemp = Choose(IntLevel, vbRed, vbBlue, vbYellow)
    Else
        SngTemp = vbBlack
    End If

    GetLevelForeColor = SngTemp
End Function

Public Function ShowFilter() As Boolean
    
End Function

Private Function GetColDataType(ColIndex As Long) As String
    Dim StrTemp As String

    With FG

        If FG.ColDataType(ColIndex) = flexDTCurrency Or FG.ColDataType(ColIndex) = flexDTDecimal Or FG.ColDataType(ColIndex) = flexDTDouble Or FG.ColDataType(ColIndex) = flexDTLong Or FG.ColDataType(ColIndex) = flexDTSingle Or FG.ColDataType(ColIndex) = flexDTLong8 Then
            StrTemp = "N"
        ElseIf FG.ColDataType(ColIndex) = flexDTDate Then
            StrTemp = "D"
        ElseIf FG.ColDataType(ColIndex) = flexDTString Or FG.ColDataType(ColIndex) = flexDTStringC Or FG.ColDataType(ColIndex) = flexDTStringW Or FG.ColDataType(ColIndex) = 130 Then
            StrTemp = "S"
        End If

    End With

    GetColDataType = StrTemp
End Function

Public Property Get sql() As String
Attribute sql.VB_ProcData.VB_Invoke_Property = "VsGroupPage"
    sql = m_SQL
End Property

Public Property Let sql(ByVal vNewValue As String)
    m_SQL = vNewValue
End Property

Public Function PrintData()
    Dim Frm As FrmViewListPrint
    RemovePointer
    Set Frm = New FrmViewListPrint
    Frm.VSPrinter1.Zoom = 100
    Frm.VSPrinter1.MarginLeft = 400
    Frm.VSPrinter1.MarginRight = 400
    Frm.VSPrinter1.StartDoc
    Frm.VSPrinter1.CurrentX = 100
    Frm.VSPrinter1.CurrentY = 100
    Frm.VSPrinter1.Text = "»«Ì  ··»—„ÃÌ« "
    Frm.VSPrinter1.CurrentX = 100
    Frm.VSPrinter1.CurrentY = 500
    Frm.VSPrinter1.RenderControl = FG.hWnd
    Frm.VSPrinter1.EndDoc
    Frm.show
End Function

Public Property Get ShowTreeGroups() As Boolean
Attribute ShowTreeGroups.VB_Description = "⁄—÷ ‘Ã—… «·„Ã„Ê⁄«  «„ ·«"
Attribute ShowTreeGroups.VB_ProcData.VB_Invoke_Property = "VsGroupPage"
Attribute ShowTreeGroups.VB_UserMemId = -520
    ShowTreeGroups = m_ShowTreeGroups
End Property

Public Property Let ShowTreeGroups(ByVal vNewValue As Boolean)
    m_ShowTreeGroups = vNewValue

    If m_ShowTreeGroups = True Then
        TrvGroups.Visible = True
        picDrag.Visible = True

        If m_iGroups > 0 Then
            LoadTreeGroups
        End If

    Else
        TrvGroups.Visible = False
        picDrag.Visible = False
    End If

    UpdateLayout False
End Property

Private Sub LoadTreeGroups()
    Dim i As Long
    Dim TreeNode As MSComctlLib.Node
    Dim FgNode As VSFlex8UCtl.VSFlexNode
    Dim LngParNodeRow As Long
    Dim SngForColor As Single

    With TrvGroups
        .Sorted = False
        .Nodes.Clear

        If Me.SetRTL = True Then
            .Nodes.Add , , "root", "«·„Ã„Ê⁄« ", "Root", "Root"
        Else
            .Nodes.Add , , "root", "Groups", "Root", "Root"
        End If

    End With

    With FG

        For i = .FixedRows To .Rows - 1

            If .IsSubtotal(i) = True Then
                Set FgNode = .GetNode(i)

                If Not FgNode.GetNode(flexNTParent) Is Nothing Then
                    LngParNodeRow = FgNode.GetNode(flexNTParent).Row
                    Set TreeNode = TrvGroups.Nodes.Add(LngParNodeRow & "G", tvwChild, i & "G", FgNode.Text, "Close", "Selected")
                Else
                    Set TreeNode = TrvGroups.Nodes.Add("root", tvwChild, i & "G", FgNode.Text, "Close", "Selected")
                End If

                TreeNode.ExpandedImage = "Open"
                SngForColor = FG.Cell(flexcpForeColor, i, 0, i, FG.Cols - 1)
                TreeNode.ForeColor = SngForColor
            End If

        Next i

    End With

    TrvGroups.Nodes("root").EnsureVisible
    TrvGroups.Nodes("root").Expanded = True
End Sub

Private Function GetFirstVisibleCol() As Long
    Dim i As Long
    Dim LngTemp As Long

    With FG

        For i = 0 To FG.Cols - 1

            If .ColHidden(i) = False Then
                LngTemp = i
                Exit For
            End If

        Next i

    End With

    GetFirstVisibleCol = LngTemp
End Function

Private Sub picDrag_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    ' start dragging with left button
    If Button <> 1 Then Exit Sub
    
    ' use trick so vertical movement is not visible
    picDrag.Height = 32000
    picDrag.top = -10000
    picDrag.Drag
End Sub

Private Sub DragDropPostion()

    If Me.SetRTL = True Then
        picDrag.left = (FG.left + FG.Width)
        picDrag.top = FG.top
        picDrag.Height = FG.Height
    
        TrvGroups.top = FG.top
        TrvGroups.Height = FG.Height
        TrvGroups.left = (picDrag.left + picDrag.Width)
        TrvGroups.Width = UserControl.ScaleWidth - (FG.Width + picDrag.Width)
    ElseIf Me.SetRTL = False Then
    
        picDrag.left = (TrvGroups.left + TrvGroups.Width)
        picDrag.top = FG.top
        picDrag.Height = FG.Height
    
        TrvGroups.top = FG.top
        TrvGroups.Height = FG.Height
        FG.left = (picDrag.left + picDrag.Width)
        FG.Width = UserControl.ScaleWidth - (TrvGroups.Width + picDrag.Width)
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SeparatorColor", picDrag.backcolor, UserControl.backcolor
    PropBag.WriteProperty "SetRTL", Me.SetRTL, True
End Sub

Public Property Get SeparatorColor() As Single
Attribute SeparatorColor.VB_Description = "·Ê‰ «·Õœ «·›«’· »Ì‰ ÃœÊ· «·»Ì«‰«  Ê‘Ã—… «·„Ã„Ê⁄« "
Attribute SeparatorColor.VB_ProcData.VB_Invoke_Property = "VsGroupPage;Appearance"
    SeparatorColor = m_SeparatorColor
End Property

Public Property Let SeparatorColor(ByVal vNewValue As Single)
    m_SeparatorColor = vNewValue
    PropertyChanged "SeparatorColor"
    picDrag.backcolor = m_SeparatorColor
End Property

Private Sub PutPointer()
    LngPointerInterval = 0

    If Me.PointerType = ArrowPointer Then
        If (ImgPointer.Visible = True And val(ImgPointer.Tag) <> 0) Then
            Timer1.Enabled = True
            ImgPointer.top = FG.RowPos(val(ImgPointer.Tag))
            ImgPointer.left = FG.Width / 2
            ImgPointer.Visible = True
            ImgPointer.ZOrder 0
        End If

    ElseIf Me.PointerType = ShapeBorderPointer Then

        If (ShpPointer.Visible = True And val(ShpPointer.Tag) <> 0) Then
            Timer1.Enabled = True
            ShpPointer.top = FG.RowPos(val(ShpPointer.Tag)) - 50
            ShpPointer.Height = FG.RowHeight(val(ShpPointer.Tag)) + 100
            ShpPointer.Width = FG.Width - 500
            ShpPointer.left = 100
            ShpPointer.Visible = True
            ShpPointer.ZOrder 0
        End If
    End If

End Sub

Private Sub RemovePointer()
    Timer1.Enabled = False

    ImgPointer.Visible = False
    ImgPointer.Tag = ""
    ShpPointer.Visible = False
    ShpPointer.Tag = ""
End Sub

Public Property Get PointerType() As PointersTypes
    PointerType = m_PointerType
End Property

Public Property Let PointerType(ByVal vNewValue As PointersTypes)
    m_PointerType = vNewValue
End Property
