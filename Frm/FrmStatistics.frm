VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStatistics 
   Caption         =   "≈Õ’«∆Ì« "
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "FrmStatistics.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleFooter 
      Height          =   795
      Left            =   210
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6480
      Width           =   6165
      _cx             =   10874
      _cy             =   1402
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
      Appearance      =   5
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
      Begin VB.Label LblComment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   555
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   5715
      End
   End
   Begin C1SizerLibCtl.C1Elastic lblSelection 
      Height          =   435
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   7695
      _cx             =   13573
      _cy             =   767
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
      CaptionPos      =   7
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
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
         ButtonImage     =   "FrmStatistics.frx":038A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   345
         Index           =   1
         Left            =   1050
         TabIndex        =   11
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ‰›Ì–"
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
         ButtonImage     =   "FrmStatistics.frx":0724
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin C1SizerLibCtl.C1Elastic lblDir 
      Height          =   765
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   7695
      _cx             =   13573
      _cy             =   1349
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
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   330
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   39213
      End
      Begin MSComCtl2.DTPicker DtpTO 
         Height          =   330
         Left            =   60
         TabIndex        =   6
         Top             =   375
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   39213
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï"
         Height          =   300
         Index           =   1
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   435
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   225
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FgDataDetails 
      Height          =   915
      Left            =   630
      TabIndex        =   3
      Top             =   3660
      Width           =   3525
      _cx             =   6218
      _cy             =   1614
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
   Begin VSFlex8UCtl.VSFlexGrid FgTree 
      Height          =   6165
      Left            =   8130
      TabIndex        =   0
      Top             =   0
      Width           =   3465
      _cx             =   6112
      _cy             =   10874
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
   Begin VSFlex8UCtl.VSFlexGrid FgData 
      Height          =   2865
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   7695
      _cx             =   13573
      _cy             =   5054
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
   Begin VB.Label picDrag 
      BackColor       =   &H00808080&
      Height          =   6105
      Left            =   7890
      TabIndex        =   2
      Top             =   30
      Width           =   90
   End
End
Attribute VB_Name = "FrmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ActionType
    GetLayout
    GetData
    GetLayoutAndData
End Enum

Private StrCurrentID As String

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

        Case 1

            If StrCurrentID = "Items1" Then
                LoadItemsSalesShow GetData
            ElseIf StrCurrentID = "Sales1" Then
                LoadSalesShow 1, GetData
            End If

    End Select

End Sub

Private Sub FgData_DragDrop(Source As Control, _
                            x As Single, _
                            Y As Single)

    ' finished resizing: resize fgTree, then fix the other controls
    If SystemOptions.UserInterface = ArabicInterface Then
        'FgTree.Width = FgTree.Width + FgTree.left + X
        FgTree.left = x
        FgTree.Width = (Me.ScaleWidth - x)
        UpdateLayout
    End If

End Sub

Private Sub FgTree_DblClick()

    With Me.FgTree

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Rowdata(.Row) = "" Or IsEmpty(.Rowdata(.Row)) Then Exit Sub
        StrCurrentID = .Rowdata(.Row)

        If .Rowdata(.Row) = "Items1" Then
            LoadItemsSalesShow GetLayout
        ElseIf .Rowdata(.Row) = "Sales1" Then
            LoadSalesShow 1, GetLayout
        ElseIf .Rowdata(.Row) = "Sales2" Then
            LoadSalesShow 1, GetLayout
        ElseIf .Rowdata(.Row) = "Customers1" Then
            'LoadCustomersShow
        End If

        '        SELECT     TOP 100 PERCENT *
        'FROM         (SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, SUM(QryTransactionsTotal.TotalAfterTax) AS TotalCus, COUNT(Transaction_ID)
        '                                              AS CountX
        '                        FROM         dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN
        '                                              dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID
        '                        GROUP BY dbo.TblCustemers.CusID, dbo.TblCustemers.CusName) DERIVEDTBL
        'ORDER BY TotalCus DESC
        'End If
    End With

End Sub

Private Sub FgTree_DragDrop(Source As Control, _
                            x As Single, _
                            Y As Single)

    If SystemOptions.UserInterface = ArabicInterface Then
        'FgTree.Width = FgTree.Width + FgTree.left + X
        FgTree.left = FgTree.left + x
        FgTree.Width = (Me.ScaleWidth - FgTree.left)
        UpdateLayout
    End If

End Sub

Private Sub Form_Load()
    Me.Width = 12000
    Me.Height = 9500
    Resize_Form Me
    InitControls
    LoadData
End Sub

Private Sub InitControls()
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo

    ' initialize tree control
    With FgTree
        
        ' structure
        .Cols = 1
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .ColAlignment(0) = flexAlignRightCenter

        If SystemOptions.UserInterface = ArabicInterface Then
            .left = (Me.ScaleWidth - .Width)
        End If

        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        .NodeClosedPicture = mdifrmmain.ImgLstTree.ListImages("Close").Picture
        .NodeOpenPicture = mdifrmmain.ImgLstTree.ListImages("OpenFolder").Picture
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .AllowSelection = False
        .Highlight = flexHighlightWithFocus
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    ' initialize list control
    With FgData

        ' structure
        .Cols = 3
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1

        ' appearance
        .BorderStyle = flexBorderNone
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .FocusRect = flexFocusNone
        .AllowUserResizing = flexResizeColumns
        .Ellipsis = flexEllipsisEnd

        ' behavior
        .AllowSelection = False
        .ExplorerBar = flexExSortShowAndMove
        .Highlight = flexHighlightAlways
        .ScrollTrack = True
        .AllowUserResizing = flexResizeColumns
        .AutoSearch = flexSearchFromCursor
        .RowHeightMin = 300
        .SortAscendingPicture = mdifrmmain.ImgLstMenuIcons.ListImages("SortASC").Picture
        .SortDescendingPicture = mdifrmmain.ImgLstMenuIcons.ListImages("SortDESC").Picture
    End With
    
    With FgDataDetails

        ' structure
        .Cols = 3
        .Rows = 1
        .FixedCols = 0
        .FixedRows = 1

        ' appearance
        .BorderStyle = flexBorderNone
        .GridLines = flexGridNone
        .BackColorBkg = .BackColor
        .SheetBorder = .BackColor
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .FocusRect = flexFocusNone
        .AllowUserResizing = flexResizeColumns
        .Ellipsis = flexEllipsisEnd

        ' behavior
        .ExplorerBar = flexExSortShowAndMove
        .Highlight = flexHighlightAlways
        .ScrollTrack = True
        .AllowUserResizing = flexResizeColumns
        .AutoSearch = flexSearchFromCursor
        .Visible = False
    End With
    
    '    'initialize label control
    With lblDir

        ' appearance
        .FontName = "Tahoma"
        .FontBold = True
        .FontSize = 14
        .BackColor = FgTree.BackColor
        .Appearance = apInsetLight
        .Caption = "‰Ê⁄ «·√Õ’«∆Ì…"
        .CaptionPos = cpRightCenter
    End With

    '
    '    ' initialize selection control
    With lblSelection

        ' appearance
        .FontName = "Tahoma"
        .FontSize = 8
        .ForeColor = vbBlue
        .CaptionPos = cpRightCenter
        .BackColor = FgData.BackColorFixed
        .Caption = " ›«’Ì· «·√Õ’«∆Ì…"

    End With
    
End Sub

Sub UpdateLayout()
    
    ' make sure this won't crash with small dimensions
    On Error Resume Next
    
    ' initialize parameters
    Dim iBorder As Single, iTreeWidth As Single, iLabelHeight As Single
    iBorder = 6 * Screen.TwipsPerPixelX
    iTreeWidth = FgTree.Width
    iLabelHeight = 750
    
    'move and position tree control
    With FgTree

        If SystemOptions.UserInterface = ArabicInterface Then
            .left = (Me.ScaleWidth - iTreeWidth)
        End If

        .Move .left, .top, iTreeWidth, Me.ScaleHeight - 100
    End With

    ' move and position drag control
    With Me.picDrag
        .Move (FgTree.left - iBorder), Me.ScaleTop, iBorder, Me.ScaleHeight
        .MousePointer = flexSizeEW ' to show resizing cursor
        .BackStyle = 0 ' transparent
        .BackColor = BackColor ' don't show this control, only the mouse cursor
    End With

    With Me.lblDir
        .left = Me.ScaleLeft + 50
        .top = Me.FgTree.top
        .Width = Me.ScaleWidth - (Me.FgTree.Width + iBorder)
        .Height = iLabelHeight
    End With

    With FgData
        lblSelection.Move lblDir.left, lblDir.Height, lblDir.Width, lblSelection.Height
    End With

    With Me.FgData
        .left = Me.lblDir.left
        .top = (Me.lblDir.Height + lblSelection.Height)
        .Width = lblDir.Width
        .Height = Me.FgTree.Height - (Me.lblDir.Height + Me.lblSelection.Height + Me.EleFooter.Height)
    End With

    With Me.EleFooter
        .left = Me.lblDir.left
        .top = (Me.FgData.top + FgData.Height)
        .Width = Me.FgData.Width
    End With
    
End Sub

Private Sub Form_Resize()
    UpdateLayout
End Sub

Private Sub picDrag_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)

    ' start dragging with left button
    If Button <> 1 Then Exit Sub
    
    ' use trick so vertical movement is not visible
    picDrag.Height = 32000
    picDrag.top = -10000
    picDrag.Drag
End Sub

Private Sub LoadData()

    With Me.FgTree
        .OutlineBar = flexOutlineBarComplete
        .AddItem "«·√Õ’«∆Ì« "
        .RowOutlineLevel(0) = 1
        .IsSubtotal(.Rows - 1) = True
    
        .AddItem "≈Õ’«∆Ì«  «·√’‰«›"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        .AddItem "„ﬁ«—‰… »Ì‰ «·√’‰«› „‰ ÕÌÀ ﬂ„Ì… «·„»Ì⁄«  Ê’«›Ï «·—»Õ"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "≈Õ’«∆Ì«  «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        .AddItem "√ﬂ»— 30 ÌÊ„ ›Ï ÕÃ„ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales1"
        
        .AddItem "√ﬂ»— 30 ÌÊ„ ›Ï ⁄œœ «·›Ê« Ì—"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales2"
        
        .AddItem "√ﬁ· 30 ÌÊ„ ›Ï ÕÃ„ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales3"
        
        .AddItem "√ﬁ· 30 ÌÊ„ ›Ï ⁄œœ «·›Ê« Ì—"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales4"
        
        .AddItem "√ﬂ»— 30 ﬁÌ„… ›« Ê—… „»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales5"
        
        .AddItem "√ﬁ· 30 ﬁÌ„… ›« Ê—… „»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Sales6"
        
        .AddItem "≈Õ’«∆Ì«  «·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
        .AddItem "„ﬁ«—‰… »Ì‰ «·⁄„·«¡ „‰ ÕÌÀ ÕÃ„ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Customers1"
        .AddItem "„ﬁ«—‰… »Ì‰ «·⁄„·«¡ „‰ ⁄œœ «·›Ê« Ì—"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Customers2"
        
        .AddItem "≈Õ’«∆Ì«  «·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
        .AddItem "≈Õ’«∆Ì«  «·„ÊŸ›Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
    End With

End Sub

Private Sub LoadItemsSalesShow(m_Action As ActionType)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim DblMaxTotal As Double

    If m_Action = GetLayout Or m_Action = GetLayoutAndData Then

        With Me.FgData
            .Clear flexClearEverywhere, flexClearEverything
            .Rows = 0
            .Cols = 0
        
            .Cols = 6
            .Rows = 1
            .FixedRows = 1
            .ColKey(0) = "Serial"
            .ColKey(1) = "ItemID"
            .ColKey(2) = "ItemCode"
            .ColKey(3) = "ItemName"
            .ColKey(4) = "Total"
            .ColKey(5) = "ProfitTotals"

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Serial")) = "„"
                .TextMatrix(0, .ColIndex("ItemID")) = "—ﬁ„ «·’‰›"
                .TextMatrix(0, .ColIndex("ItemCode")) = "ﬂÊœ «·’‰›"
                .TextMatrix(0, .ColIndex("ItemName")) = "«”„ «·’‰›"
                .TextMatrix(0, .ColIndex("Total")) = "≈Ã„«·Ï «·„»Ì⁄« "
                .TextMatrix(0, .ColIndex("ProfitTotals")) = "≈Ã„«·Ï ’«›Ï «·—»Õ"

                For i = 0 To .Cols - 1
                    .ColAlignment(i) = flexAlignRightCenter
                    .FixedAlignment(i) = flexAlignRightCenter
                Next i

            Else
                .TextMatrix(0, .ColIndex("Serial")) = "Serial"
                .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
                .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
                .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
                .TextMatrix(0, .ColIndex("Total")) = "Totals"

                For i = 0 To .Cols - 1
                    .ColAlignment(i) = flexAlignLeftCenter
                    .FixedAlignment(i) = flexAlignLeftCenter
                Next i

            End If

            .AutoSize 0, .Cols - 1, False
        End With

    ElseIf m_Action = GetData Or m_Action = GetLayoutAndData Then

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT ItemID, ItemName, ItemCode, Total"
            StrSQL = StrSQL + " FROM QryItemsTransactionsTotals(2,0," ''1/31/2005','1/1/2008') "

            If Not IsNull(Me.DTPFrom.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPFrom.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/1900#, True)
            End If

            StrSQL = StrSQL + ","

            If Not IsNull(Me.DTPTo.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPTo.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/2079#, True)
            End If

            StrSQL = StrSQL + ")"
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT ItemID, ItemName, ItemCode, Total"
            StrSQL = StrSQL + " FROM QryItemsTransactionsTotals(2,0," ''1/31/2005','1/1/2008') "

            If Not IsNull(Me.DTPFrom.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPFrom.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/1900#, True)
            End If

            StrSQL = StrSQL + ","

            If Not IsNull(Me.DTPTo.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPTo.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/2079#, True)
            End If

            StrSQL = StrSQL + ")"
        End If
    
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            With Me.FgData
                .Rows = .FixedRows + rs.RecordCount

                For i = 1 To rs.RecordCount
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
                    rs.MoveNext
                Next i

                DblMaxTotal = .Aggregate(flexSTMax, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))

                If DblMaxTotal <> 0 Then
                    .Cell(flexcpFloodColor, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total")) = &HC0&

                    For i = .FixedRows To .Rows - 1
                        .Cell(flexcpFloodPercent, i, .ColIndex("Total")) = 100 * val(.TextMatrix(i, .ColIndex("Total"))) / DblMaxTotal
                    Next i

                End If

                .Refresh
                .AutoSize 0, .Cols - 1, False
            End With

        End If

        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub LoadSalesShow(Index As Integer, _
                          m_Action As ActionType)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim DblMaxTotal As Double
    Dim DblTotalPeriod As Double
    Dim DblPercent As Double
    Dim XFont As IFontDisp

    'Index=1 ........√ﬂ»— 30 ÌÊ„ ›Ï ÕÃ„ «·„»Ì⁄« 
    'Index=2 ........√ﬂ»— 30 ÌÊ„ ›Ï ⁄œœ «·›Ê« Ì—
    'Index=3 ........
    'Index=4 ........
    If m_Action = GetLayout Or m_Action = GetLayoutAndData Then

        With Me.FgData
            .Clear flexClearEverywhere, flexClearEverything
            .Rows = 0
            .Cols = 0
            .Cols = 5
            .Rows = 1
            .FixedRows = 1
            .ColKey(0) = "Serial"
            .ColKey(1) = "Transaction_Date"
            .ColDataType(1) = flexDTDate
            .ColKey(2) = "SumX"
            .ColKey(3) = "CountX"
            .ColKey(4) = "Percent"

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(0, .ColIndex("Serial")) = "„"
                .TextMatrix(0, .ColIndex("Transaction_Date")) = " «—ÌŒ «·ÌÊ„"
                .TextMatrix(0, .ColIndex("SumX")) = "≈Ã„«·Ï „»Ì⁄«  «·ÌÊ„"
                .TextMatrix(0, .ColIndex("CountX")) = "⁄œœ ›Ê« Ì— «·ÌÊ„"
                .TextMatrix(0, .ColIndex("Percent")) = "‰”»… „»Ì⁄«  «·ÌÊ„ ≈·Ï „»Ì⁄«  «·› —…"

                For i = 0 To .Cols - 1
                    .ColAlignment(i) = flexAlignRightCenter
                    .FixedAlignment(i) = flexAlignRightCenter
                Next i

            Else
                .TextMatrix(0, .ColIndex("Serial")) = "Serial"
                .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
                .TextMatrix(0, .ColIndex("SumX")) = "Day Total Sales"
                .TextMatrix(0, .ColIndex("CountX")) = "Day Total Count Inv.s"
                .TextMatrix(0, .ColIndex("Percent")) = "Day Sales Percent"

                For i = 0 To .Cols - 1
                    .ColAlignment(i) = flexAlignLeftCenter
                    .FixedAlignment(i) = flexAlignLeftCenter
                Next i

            End If

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    If m_Action = GetData Or m_Action = GetLayoutAndData Then
        '------------------------------------------------------------------------------
        DblTotalPeriod = GetTransactionTotalPeriod(2, Me.DTPFrom.value, Me.DTPTo.value)

        '------------------------------------------------------------------------------
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT     TOP 30 * FROM "
            StrSQL = StrSQL + " ("
            StrSQL = StrSQL + " SELECT Transaction_Date, SUM(TotalAfterTax) AS SumX," & "COUNT(Transaction_ID) AS CountX "
            StrSQL = StrSQL + " FROM  dbo.QryTransactionsTotal() QryTransactionsTotal "
            StrSQL = StrSQL + " Where QryTransactionsTotal.Transaction_Type=2"

            If Not IsNull(Me.DTPFrom.value) Then
                StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(Me.DTPFrom.value, True)
            End If

            If Not IsNull(Me.DTPTo.value) Then
                StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(Me.DTPTo.value, True)
            End If

            StrSQL = StrSQL + " GROUP BY Transaction_Date) DERIVEDTBL "
            StrSQL = StrSQL + " ORDER BY SumX DESC "
       
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT ItemID, ItemName, ItemCode, Total"
            StrSQL = StrSQL + " FROM QryItemsTransactionsTotals(2,0," ''1/31/2005','1/1/2008') "

            If Not IsNull(Me.DTPFrom.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPFrom.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/1900#, True)
            End If

            StrSQL = StrSQL + ","

            If Not IsNull(Me.DTPTo.value) Then
                StrSQL = StrSQL + SQLDate(Me.DTPTo.value, True)
            Else
                StrSQL = StrSQL + SQLDate(#1/1/2079#, True)
            End If

            StrSQL = StrSQL + ")"
        End If
    
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then

            With Me.FgData
                .Rows = .FixedRows + rs.RecordCount

                For i = 1 To rs.RecordCount
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", DisplayDate(rs("Transaction_Date").value))
                    .TextMatrix(i, .ColIndex("SumX")) = IIf(IsNull(rs("SumX").value), "", rs("SumX").value)
                    .TextMatrix(i, .ColIndex("CountX")) = IIf(IsNull(rs("CountX").value), "", rs("CountX").value)
                    DblPercent = (val(.TextMatrix(i, .ColIndex("SumX"))) * 100) / DblTotalPeriod
                    .TextMatrix(i, .ColIndex("Percent")) = Format(DblPercent, SystemOptions.SysDefCurrencyForamt)
                    rs.MoveNext
                Next i

                ModFgLib.DrawFloodProgress Me.FgData, FgData.ColIndex("SumX"), &H80FF&
                ModFgLib.DrawFloodProgress Me.FgData, FgData.ColIndex("CountX"), &HFF8080
                ModFgLib.DrawFloodProgress Me.FgData, FgData.ColIndex("Percent"), &HFF8080, DblTotalPeriod, FgData.ColIndex("SumX")
                Set XFont = Me.Font
                XFont.name = "Tahoma"
                XFont.Size = 10
                XFont.Charset = 178
                .Cell(flexcpFont, .FixedRows, 0, .Rows - 1, .Cols - 1) = XFont
                .Cell(flexcpFontBold, .FixedRows, 0, .Rows - 1, .Cols - 1) = True
                .Refresh
                .AutoSize 0, .Cols - 1, False, 300
            End With

        End If

        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub LoadCustomersShow(Index As Integer, _
                              m_Action As ActionType)

End Sub

