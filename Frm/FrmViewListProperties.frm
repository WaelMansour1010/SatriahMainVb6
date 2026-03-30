VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmViewListProperties 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " ÕœÌœ Œ’«∆’ «·⁄—÷"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7035
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4890
      _cx             =   8625
      _cy             =   12409
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
      _GridInfo       =   $"FrmViewListProperties.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   525
         Left            =   15
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6495
         Width           =   4860
         _cx             =   8573
         _cy             =   926
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
         AutoSizeChildren=   0
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   375
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            Caption         =   "Õ›Ÿ"
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
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   675
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   4860
         _cx             =   8573
         _cy             =   1191
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
         Caption         =   "ÿ—Ìﬁ… Ê‰Ÿ«„ «·⁄—÷"
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
         Begin VB.OptionButton optAlpha 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ÿ«„ «·⁄—÷ ﬂ — Ì» «»ÃœÏ"
            Height          =   255
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   270
            Width           =   2445
         End
         Begin VB.OptionButton optCat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ÿ«„ «·⁄—÷ ﬂ„Ã„Ê⁄« "
            Height          =   255
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   2145
         End
         Begin MSComDlg.CommonDialog cmDlg 
            Left            =   2520
            Top             =   150
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image imgColorPick 
            Height          =   240
            Left            =   270
            Picture         =   "FrmViewListProperties.frx":007B
            Top             =   180
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgFontPick 
            Height          =   240
            Left            =   660
            Picture         =   "FrmViewListProperties.frx":01C5
            Top             =   120
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5775
         Left            =   15
         TabIndex        =   1
         Top             =   705
         Width           =   4860
         _cx             =   8572
         _cy             =   10186
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
         SubtotalPosition=   1
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
      End
   End
End
Attribute VB_Name = "FrmViewListProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' property type
Private Type Property_Type
    category As String
    name As String
    ptype As Integer
    value As Variant
End Type

Private Enum PropertyType_Type
    ptNil
    ptcolor
    ptBool
    ptFont
    ptFontName
    ptValue
End Enum

' property vector
Dim g_Props(15) As Property_Type

' string to hold font list
Dim g_FontList As String

' API Declarations (for OwnerDraw cells)
Private Declare Function SetBkColor _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetSysColor _
                Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function ExtTextOut _
                Lib "gdi32" _
                Alias "ExtTextOutA" (ByVal hDC As Long, _
                                     ByVal x As Long, _
                                     ByVal Y As Long, _
                                     ByVal wOptions As Long, _
                                     lpRect As RECT, _
                                     ByVal lpString As String, _
                                     ByVal nCount As Long, _
                                     lpDx As Long) As Long

Private Declare Function GetStockObject _
                Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function FrameRect _
                Lib "user32" (ByVal hDC As Long, _
                              lpRect As RECT, _
                              ByVal hBrush As Long) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const ETO_OPAQUE = 2

Private Const BLACK_BRUSH = 4

Private Sub DisplayCategorized()
    
    ' freeze to avoid flicker
    FG.Redraw = flexRDNone
    
    ' remove any existing subtotals (groups)
    FG.Subtotal flexSTClear
    
    ' sort by category, then by property name
    FG.Select 1, 1, 1, 2
    FG.Sort = flexSortStringAscending
    
    ' add subtotals (groups) by category (col 1)
    FG.Subtotal flexSTNone, 1, , , &HE2E9E9, &H40&, True

    ' show outline column
    FG.ColHidden(0) = False
    
    ' to look nice
    FG.GridLines = flexGridFlatVert
    
    ' reset display
    FG.TopRow = 1
    FG.Select 2, FG.Cols - 1
    FG.Redraw = flexRDBuffered

End Sub

Private Sub DisplayAlphabetic()
    
    ' freeze to avoid flicker
    FG.Redraw = flexRDNone
    
    ' remove any existing subtotals (groups)
    FG.Subtotal flexSTClear
    
    ' sort by property name
    FG.Col = 2
    FG.Sort = flexSortStringAscending
    
    ' hide outline column
    FG.ColHidden(0) = True
    
    ' to look nice
    FG.GridLines = flexGridFlat
    
    ' reset display
    FG.TopRow = 1
    FG.Select 1, FG.Cols - 1
    FG.Redraw = flexRDBuffered

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If FG.ComboList <> "" Then
        If val(FG.ComboData) <> 0 Then
            FG.Cell(flexcpData, Row, Col) = FG.ComboData
            'FG.TextMatrix(Row, Col) = ""
        End If
    End If

End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)
    
    ' we can't edit total rows or label columns
    If FG.IsSubtotal(Row) Or Col <> FG.Cols - 1 Then
        Cancel = True
        Exit Sub
    End If
    
    ' assume regular editing
    FG.ComboList = ""
    
    ' setup to edit based on property type
    Select Case g_Props(FG.Rowdata(Row)).ptype
    
            ' font name gets a combo
        Case ptFontName
            FG.ComboList = g_FontList

            ' fonts get a pick button
        Case ptFont
            FG.ComboList = "..."
            FG.CellButtonPicture = imgFontPick
    
            ' colors get a different pick button
        Case ptcolor
            FG.ComboList = "..."
            FG.CellButtonPicture = imgColorPick
    
            ' booleans get a pick list
        Case ptBool
            FG.ComboList = "#1;ÃœÌœ|#2;„” ⁄„·"
    End Select

    ' use automatic double-click for editing text, manual for lists
    If Len(FG.ComboList) Then
        FG.Editable = flexEDKbd
    Else
        FG.Editable = flexEDKbdMouse
    End If

End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, _
                                  ByVal OldCol As Long, _
                                  ByVal NewRow As Long, _
                                  ByVal NewCol As Long, _
                                  Cancel As Boolean)

    ' user can select only the last column
    With FG

        If .Redraw <> flexRDNone And NewCol <> .Cols - 1 Then
            Cancel = True
            .Select NewRow, .Cols - 1
        End If

    End With
    
End Sub

Private Sub fg_BeforeUserResize(ByVal Row As Long, _
                                ByVal Col As Long, _
                                Cancel As Boolean)

    ' don't resize outline column
    If Col = 0 Then Cancel = True
    
End Sub

Private Sub Fg_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)
    cmDlg.CancelError = False
    
    ' clicked the button to edit a color?
    If g_Props(FG.Rowdata(Row)).ptype = ptcolor Then
    
        ' position palette form below current cell
        FrmPalette.Move left + FG.left + FG.ColPos(Col) + (Width - ScaleWidth), top + FG.top + FG.RowPos(Row) + FG.RowHeight(Row) + (Height - ScaleHeight)

        ' show the palette
        ' we show the palette as a modeless dialog and wait until the user is done with it, either
        ' by clicking a color, pressing ESC, or just by activating some other window.
        ' alternatively, we could show it modally and then we could remove the While statement.
        '
        FrmPalette.Show vbModal
        While FrmPalette.Visible

            DoEvents
        Wend
        
        ' if the user picked a value, it's in the Tag property
        If FrmPalette.Tag <> "" Then FG.TextMatrix(Row, FG.Cols - 1) = FormatColor(FrmPalette.Tag)
        Unload FrmPalette
        
        ' we could use the common color dialog instead, but it has no support for system colors
        'cmDlg.Flags = cdlCCRGBInit
        'cmDlg.Color = Val(fg.TextMatrix(Row, fg.Cols - 1))
        'cmDlg.ShowColor
        'fg.TextMatrix(Row, fg.Cols - 1) = FormatColor(cmDlg.Color)
        
        ' clicked the button to edit a font?
    ElseIf g_Props(FG.Rowdata(Row)).ptype = ptFont Then
        cmDlg.Flags = cdlCFBoth Or cdlCCRGBInit Or cdlCFEffects
        cmDlg.FontName = FG.Cell(flexcpFontName, Row, FG.Cols - 1)
        cmDlg.FontBold = FG.Cell(flexcpFontBold, Row, FG.Cols - 1)
        cmDlg.FontItalic = FG.Cell(flexcpFontItalic, Row, FG.Cols - 1)
        cmDlg.FontSize = FG.Cell(flexcpFontSize, Row, FG.Cols - 1)
        cmDlg.FontUnderline = FG.Cell(flexcpFontUnderline, Row, FG.Cols - 1)
        cmDlg.FontStrikethru = FG.Cell(flexcpFontStrikethru, Row, FG.Cols - 1)
        cmDlg.color = FG.Cell(flexcpForeColor, Row, FG.Cols - 1)
        cmDlg.ShowFont
        FG.TextMatrix(Row, FG.Cols - 1) = cmDlg.FontName
        
        ' format cell to show font
        FG.Cell(flexcpFontName, Row, FG.Cols - 1) = cmDlg.FontName
        FG.Cell(flexcpFontBold, Row, FG.Cols - 1) = cmDlg.FontBold
        FG.Cell(flexcpFontItalic, Row, FG.Cols - 1) = cmDlg.FontItalic
        FG.Cell(flexcpFontSize, Row, FG.Cols - 1) = cmDlg.FontSize
        FG.Cell(flexcpFontUnderline, Row, FG.Cols - 1) = cmDlg.FontUnderline
        FG.Cell(flexcpFontStrikethru, Row, FG.Cols - 1) = cmDlg.FontStrikethru
        FG.Cell(flexcpForeColor, Row, FG.Cols - 1) = cmDlg.color
    
    End If
    
End Sub

Private Sub Fg_DblClick()

    ' double-clicking on a group collapses/expands it
    Dim r%
    r = FG.MouseRow

    If r > -1 Then
        If FG.IsSubtotal(r) Then
            If FG.IsCollapsed(r) = flexOutlineCollapsed Then
                FG.IsCollapsed(r) = flexOutlineExpanded
            Else
                FG.IsCollapsed(r) = flexOutlineCollapsed
            End If

            FG.Tag = ""
            
            ' double-clicking on regular cells edits them
        Else
            FG.Tag = "*"
            FG.EditCell
        End If
    End If

End Sub

Private Sub fg_DrawCell(ByVal hDC As Long, _
                        ByVal Row As Long, _
                        ByVal Col As Long, _
                        ByVal left As Long, _
                        ByVal top As Long, _
                        ByVal right As Long, _
                        ByVal bottom As Long, _
                        done As Boolean)

    ' only need to custom draw color selection cells
    If Col <> 3 Then Exit Sub
    If g_Props(FG.Rowdata(Row)).ptype <> ptcolor Then Exit Sub
    
    ' build color rectangle
    Dim rc As RECT
    rc.left = left + 2
    rc.right = left + 15
    rc.top = top + 2
    rc.bottom = bottom - 3
    
    ' translate color
    Dim clr&
    clr = val(FG.TextMatrix(Row, FG.Cols - 1))

    If (clr And &H80000000) Then
        clr = GetSysColor(clr And &HFF)
    End If
    
    ' paint rectangle
    clr = SetBkColor(hDC, clr)
    ExtTextOut hDC, 0, 0, ETO_OPAQUE, rc, 0, 0, 0
    SetBkColor hDC, clr
    
    ' frame rectangle
    FrameRect hDC, rc, GetStockObject(BLACK_BRUSH)
    
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)
    Dim r%
    
    ' special handling for cursor keys
    Select Case KeyCode

            ' collapse/expand with cursor keys
        Case vbKeyLeft, vbKeyHome

            If FG.IsSubtotal(FG.Row) Then FG.IsCollapsed(FG.Row) = flexOutlineCollapsed
            If FG.Col <> FG.Cols - 1 Then FG.Col = FG.Cols - 1
            KeyCode = 0

        Case vbKeyRight, vbKeyEnd

            If FG.IsSubtotal(FG.Row) Then FG.IsCollapsed(FG.Row) = flexOutlineExpanded
            If FG.Col <> FG.Cols - 1 Then FG.Col = FG.Cols - 1
            KeyCode = 0
                        
            ' when pushing control+ASCII, look for property
        Case Else

            If Shift >= 2 And KeyCode >= Asc("A") And KeyCode <= Asc("Z") Then
            
                ' look From current row down to bottom
                For r = FG.Row + 1 To FG.Rows - 1

                    If Not FG.RowHidden(r) Then
                        If UCase(left(FG.TextMatrix(r, FG.Cols - 2), 1)) = Chr(KeyCode) Then
                            FG.Select r, FG.Cols - 1
                            FG.ShowCell r, FG.Cols - 1
                            KeyCode = 0
                            Exit For
                        End If
                    End If

                Next
                
                ' not found, so look From top down to current - 1
                For r = FG.FixedRows To FG.Row - 1

                    If Not FG.RowHidden(r) Then
                        If UCase(left(FG.TextMatrix(r, FG.Cols - 2), 1)) = Chr(KeyCode) Then
                            FG.Select r, FG.Cols - 1
                            FG.ShowCell r, FG.Cols - 1
                            KeyCode = 0
                            Exit For
                        End If
                    End If

                Next

            End If

    End Select
    
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    ' if this is a list, double-clicking selects the next item
    If Len(FG.Tag) > 0 And Len(FG.ComboList) > 0 And FG.ComboList <> "..." Then
        If ComboNext(Row, Col) Then Cancel = True
    End If

    FG.Tag = ""
    
End Sub

Private Function ComboNext(Row&, _
                           Col&) As Boolean

    ' get current list, trim combo pipe if any
    Dim s$, i%
    s = FG.ComboList

    If left(s, 1) = "|" Then s = Mid(s, 2)
    
    ' look for current text in list, fail if not found
    i = InStr(s, FG.TextMatrix(Row, Col))

    If i <= 0 Then Exit Function
    
    ' look for next choice
    i = InStr(i, s, "|")

    If (i > 0) Then s = Mid(s, i + 1)
    
    ' trim excess
    i = InStr(s, "|")

    If i > 0 Then s = left(s, i - 1)
    
    ' set new entry
    FG.TextMatrix(Row, Col) = s
    ComboNext = True
    
End Function

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion
    SetFg
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ISButton1_Click()
    SaveOptions
End Sub

Private Sub optAlpha_Click()
    DisplayAlphabetic
End Sub

Private Sub optCat_Click()
    DisplayCategorized
End Sub

Private Function FormatColor(clr As Long) As String
    
    ' translate value into fixed-length hex
    Dim s$
    s = Hex(clr)

    If Len(s) < 8 Then s = String(8 - Len(s), "0") & s

    ' prepend 'H' and some spaces to fit owner-drawn color box
    FormatColor = "     &H" & s & "&"

End Function

Private Sub InitPropertyList()
    Dim SngTempColorValue As Long
    Dim StrTempTextValue As String
    Dim BolValue As Boolean
    Dim i As Long

    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·√Ê·"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·À«‰Ï"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·À«·À"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·—«»⁄"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·Œ«„”"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·”«œ”"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ «·„” ÊÏ «·”«»⁄"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = vbBlack
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·√Ê·"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = CStr("0" & i) & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·À«‰Ï"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = i & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·À«·À"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(i <= 9, CStr("0" & i), i), 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = i & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·—«»⁄"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & i, 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = i & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·Œ«„”"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & i, 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = i & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·”«œ”"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & i, 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "«·‘ﬂ· «·⁄«„"
    g_Props(i).ptype = ptcolor
    g_Props(i).name = i & " ·Ê‰ Œ·›Ì… «·„” ÊÏ «·”«»⁄"
    SngTempColorValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & i, 0))

    If SngTempColorValue = 0 Then
        SngTempColorValue = &H8000000F
    End If

    g_Props(i).value = FormatColor(SngTempColorValue)
    '-------------------------------------------------------------------
    i = i + 1
    g_Props(i).category = "„’ÿ·Õ«  Ê„”„Ì« "
    g_Props(i).ptype = ptBool
    g_Props(i).name = "⁄—÷ ﬂ·„… ≈Ã„«·Ï"
    BolValue = val(GetSetting(SystemOptions.SysRegsAppPath, "ViewListSetting\OtherSetting", "Setting" & i, 0))

    If BolValue = False Then
        g_Props(i).value = 0
    Else
        g_Props(i).value = 1
    End If
                
End Sub

Private Sub SaveOptions()
    Dim i As Long
    Dim LngPropNumber As Long

    'BolValue = Val(GetSetting(SystemOptions.SysRegsAppPath, _
     "ViewListSetting\OtherSetting", "Setting" & I, 0))
    
    If Me.optAlpha.value = True Then

        For i = 1 To Me.FG.Rows - 1

            If Not IsEmpty(FG.Rowdata(i)) Then
                LngPropNumber = val(FG.Rowdata(i))

                If LngPropNumber >= 1 And LngPropNumber <= 14 Then
                    SaveSetting SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(LngPropNumber <= 9, CStr("0" & LngPropNumber), LngPropNumber), val(Me.FG.TextMatrix(i, 3))
                Else
                    SaveSetting SystemOptions.SysRegsAppPath, "ViewListSetting\OtherSetting", "Setting" & i, (Me.FG.TextMatrix(i, 3))
                End If
            End If

        Next i

    ElseIf optCat.value = True Then

        For i = 2 To Me.FG.Rows - 1

            If Not IsEmpty(FG.Rowdata(i)) Then
                LngPropNumber = val(FG.Rowdata(i))

                If LngPropNumber >= 1 And LngPropNumber <= 14 Then
                    SaveSetting SystemOptions.SysRegsAppPath, "ViewListSetting\ColorSetting", "SettingColor" & IIf(LngPropNumber <= 9, CStr("0" & LngPropNumber), LngPropNumber), val(Me.FG.TextMatrix(i, 3))
                Else
                    SaveSetting SystemOptions.SysRegsAppPath, "ViewListSetting\OtherSetting", "Setting" & i, (Me.FG.TextMatrix(i, 3))
                End If
            End If

        Next i

    End If

    Unload Me
End Sub

Private Sub SetFg()
    Dim i As Long

    ' build font list to display in combo
    For i = 0 To Screen.FontCount - 1
        g_FontList = g_FontList & "|" & Screen.Fonts(i)
    Next

    ' initialize property list
    InitPropertyList
    ' initialize control
    FG.Rows = 1                                 ' start empty
    FG.Cols = 4

    If SystemOptions.UserInterface = ArabicInterface Then
        FG.RightToLeft = True
        FG.ColAlignment(-1) = flexAlignRightCenter
        FG.FixedAlignment(-1) = flexAlignRightCenter
    Else
        FG.RightToLeft = False
        FG.ColAlignment(-1) = flexAlignLeftCenter
        FG.FixedAlignment(-1) = flexAlignLeftCenter
    End If

    ' outline, category, property, value
    FG.TextMatrix(0, 2) = "«·Œ«’Ì… «Ê «·„Ì“…"     ' column titles
    FG.TextMatrix(0, 3) = "«·ﬁÌ„… «·Õ«·Ì… ·Â«"    ' column titles
    FG.Editable = flexEDKbd                     ' double-clicks start editing
    FG.OwnerDraw = flexODOver                   ' use ownerdraw to show colors
    FG.OutlineCol = 0                           ' set outline column properties
    FG.ColWidth(0) = 230                        ' narrow outline column
    FG.OutlineBar = flexOutlineBarSymbolsLeaf   ' no tree, just symbols
    '    fg.ColHidden(1) = True                     ' hide categories
    FG.ColWidth(1) = 0                          ' hide categories
    FG.MergeCells = flexMergeSpill              ' allow categories to spill into property column
    'Fg.ColAlignment(-1) = flexAlignLeftTop      ' align all to left
    FG.AllowSelection = False                   ' select a single cell at a time
    FG.AllowUserResizing = flexResizeColumns    ' give user freedom
    FG.ScrollTrack = True                       ' scroll as the user drags the scroll thumb
    FG.FixedCols = 0                            ' to look nice
    FG.ExtendLastCol = True
    FG.Ellipsis = flexEllipsisEnd
    FG.BackColorBkg = FG.GridColor
    FG.Highlight = flexHighlightNever
    FG.RowHeightMin = 300

    'populate control
    For i = 1 To UBound(g_Props)
        FG.AddItem vbTab & g_Props(i).category & vbTab & g_Props(i).name & vbTab & g_Props(i).value
        FG.Rowdata(i) = i ' keep property index because we will be sorting this
    Next

    'do an autosize on property names
    FG.AutoSize FG.Cols - 2, , , 300
    ' initialize display
    DisplayCategorized

End Sub
