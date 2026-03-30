VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ItemLOtNO 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ÕœÌœ —ﬁ„ «··Êÿ"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin ImpulseButton.ISButton CmdOk 
      Default         =   -1  'True
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Top             =   2730
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   855
      Left            =   7230
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3870
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   9
      Top             =   2730
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   6360
      TabIndex        =   10
      Top             =   840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   97320961
      CurrentDate     =   38784
   End
   Begin VSFlex8Ctl.VSFlexGrid QtyGrid 
      Height          =   1635
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   5505
      _cx             =   9710
      _cy             =   2884
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   100
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ItemLotNO.frx":0000
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
      ExplorerBar     =   3
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
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«‰ «Ã"
      Height          =   255
      Index           =   6
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   660
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   660
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›: "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”ÿ—: "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "ItemLOtNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FG As VSFlex8UCtl.vsFlexGrid
Public WithEvents TxtLotNo As TextBox
Attribute TxtLotNo.VB_VarHelpID = -1

Public LngRow As Long
Public Indxx As Integer
Public LngCol As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim ExpiryDate As Date
    Dim Askinterval As String
    If Indxx = 0 Then
    If Not FG Is Nothing Then
       ' FG.TextMatrix(LngRow, LngCol) = (QtyGrid.TextMatrix(QtyGrid.Row, 1))    'Trim$(Me.TxtComment.text)

                    If Me.FG.ColIndex("ProductionDate") <> -1 Then
             
                        FG.TextMatrix(LngRow, FG.ColIndex("ProductionDate")) = (QtyGrid.TextMatrix(QtyGrid.Row, 2))  ' ExpiryDate
                    End If
            
                    If Me.FG.ColIndex("ExpiryDate") <> -1 Then
             
                        FG.TextMatrix(LngRow, FG.ColIndex("ExpiryDate")) = (QtyGrid.TextMatrix(QtyGrid.Row, 3))  ' ExpiryDate
                    End If
                        If Me.FG.ColIndex("LotNO") <> -1 Then
             
                        FG.TextMatrix(LngRow, FG.ColIndex("LotNO")) = (QtyGrid.TextMatrix(QtyGrid.Row, 1))  ' ExpiryDate
                    End If

        Unload Me
    End If
   ElseIf Indxx = 1 Then
   frmsalebill.TxtLotNo.Text = (QtyGrid.TextMatrix(QtyGrid.Row, 1))
    Unload Me
   frmsalebill.TxtQuantity.SetFocus
End If
End Sub

Public Sub FillGridWithData()

End Sub

Private Sub Form_Load()
On Error Resume Next
    CenterForm Me

    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    DoEvents
    DoEvents
    
    With QtyGrid
         .Col = .ColIndex("LotNO")
                             .Row = 1
                             .ShowCell 1, .ColIndex("LotNO")
                             
                             .SetFocus
End With
'QtyGrid.SetFocus

End Sub

Function ChangeLang()
    Me.Caption = "Lot NO"
    lbl(0).Caption = "Line"
    lbl(1).Caption = "Code"
    lbl(2).Caption = "Name"

    With QtyGrid
        .TextMatrix(0, .ColIndex("LotNO")) = "Lot NO"
        .TextMatrix(0, .ColIndex("ProductionDate")) = "ProductionDate"
        .TextMatrix(0, .ColIndex("ExpiryDate")) = "ExpiryDate"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
    End With

    CmdOk.Caption = "Save"
    CmdCancel.Caption = "Close"

End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub QtyGrid_DblClick()
    CmdOk_Click
End Sub
