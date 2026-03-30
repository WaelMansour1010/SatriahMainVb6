VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOlapShow 
   Caption         =   "⁄—÷  Õ·Ì· «·»Ì«‰« "
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   Icon            =   "FrmOlapShow.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   10260
   Begin C1SizerLibCtl.C1Elastic lblSelection 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
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
      BackColor       =   -2147483633
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
         TabIndex        =   1
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmOlapShow.frx":038A
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   345
         Index           =   1
         Left            =   1050
         TabIndex        =   2
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
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
         ButtonImage     =   "FrmOlapShow.frx":0724
         DrawFocusRectangle=   0   'False
      End
   End
   Begin C1SizerLibCtl.C1Elastic lblDir 
      Height          =   765
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
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
      BackColor       =   -2147483633
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         Caption         =   "„‰"
         Height          =   300
         Index           =   0
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   225
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï"
         Height          =   300
         Index           =   1
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   435
         Width           =   285
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid FgTree 
      Height          =   6165
      Left            =   7830
      TabIndex        =   8
      Top             =   0
      Width           =   2775
      _cx             =   4895
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
   Begin VB.Label picDrag 
      BackColor       =   &H00808080&
      Height          =   6105
      Left            =   7710
      TabIndex        =   9
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "FrmOlapShow"
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
                'LoadItemsSalesShow GetData
            End If

    End Select

End Sub

Private Sub DCube1_DragDrop(Source As Control, _
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
            'LoadItemsSalesShow GetLayout
        End If

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
    Me.Width = 10845
    Me.Height = 7920
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
        .Ellipsis = flexEllipsisEnd
        
        ' behavior
        .AllowSelection = False
        .Highlight = flexHighlightWithFocus
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    ' initialize list control
    '    With DCube1
    '
    '
    '    End With
    
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
        .BackColor = Me.BackColor
        .Caption = " ›«’Ì· «·√Õ’«∆Ì…"

    End With
    
End Sub

Sub UpdateLayout()
    
    ' make sure this won't crash with small dimensions
    On Error Resume Next
    
    ' initialize parameters
    Dim iBorder%, iTreeWidth%, iLabelHeight%
    iBorder = 6 * Screen.TwipsPerPixelX
    iTreeWidth = FgTree.Width
    iLabelHeight = 750
    
    ' move and position tree control
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

    '    With DCube1
    '        lblSelection.Move lblDir.left, lblDir.Height, lblDir.Width, lblSelection.Height
    '    End With
    '    With Me.DCube1
    ''        .left = Me.lblDir.left
    ''        .top = (Me.lblDir.Height + lblSelection.Height)
    ''        .Width = lblDir.Width
    ''        .Height = Me.FgTree.Height - (Me.lblDir.Height + Me.lblSelection.Height)
    '        .Move Me.lblDir.left, (Me.lblDir.Height + lblSelection.Height), lblDir.Width, Me.FgTree.Height - (Me.lblDir.Height + Me.lblSelection.Height)
    '        .Refresh
    '
    '    End With
    ' move and position list
    
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
    
        .AddItem "≈Õ’«∆Ì«  «·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
        .AddItem "„ﬁ«—‰… »Ì‰ «·⁄„·«¡ „‰ ÕÌÀ «·‘—«¡ «·‰ﬁœÏ"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        
        .AddItem "„ﬁ«—‰… »Ì‰ «·⁄„·«¡ „‰ ÕÌÀ ”œ«œ «·ﬁÌ„ «·„«·Ì… «·√Ã·…"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = False
        
        .AddItem "≈Õ’«∆Ì«  «·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
        .AddItem "≈Õ’«∆Ì«  «·„ÊŸ›Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
    
    End With

End Sub

