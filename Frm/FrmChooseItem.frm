VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmChooseItem 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "≈Œ Ì«— ’‰›"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "FrmChooseItem.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _cx             =   12726
      _cy             =   5265
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmChooseItem.frx":058A
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
   Begin ImpulseButton.ISButton ISBXPBtnOK 
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   3630
      Width           =   915
      _ExtentX        =   1614
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
      BackStyle       =   0
      ButtonImage     =   "FrmChooseItem.frx":0656
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton ISBXPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   3630
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
      BackStyle       =   0
      ButtonImage     =   "FrmChooseItem.frx":09F0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   345
      Index           =   2
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3630
      Width           =   465
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·√’‰«›:-"
      Height          =   345
      Index           =   1
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   465
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   7230
      X2              =   0
      Y1              =   3540
      Y2              =   3540
   End
End
Attribute VB_Name = "FrmChooseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_UserCanceld As Boolean

Public LngChooseItemID As Long

Private Sub Fg_DblClick()
    ISBXPBtnOK_Click
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)

    If KeyCode = vbKeyReturn Then
        ISBXPBtnOK_Click
    End If

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    Set FG.WallPaper = GrdBack.Picture
    Msg = "·ﬁœ ÊÃœ «·»—‰«„Ã .. «ﬂÀ— „‰ ’‰› Ì ‘«»ÂÊ‰ ›Ï ‰›” «·”Ì—Ì«·"
    Msg = Msg & Chr(13) & "›»—Ã«¡  ÕœÌœ «·’‰› «·„—«œ"
    Me.lbl(0).Caption = Msg
    CenterForm Me

    FormPostion Me, GetPostion
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ISBXPBtnCancel_Click()
    Me.UserCanceld = True
    Me.Hide
End Sub

Public Property Get UserCanceld() As Boolean
    UserCanceld = m_UserCanceld
End Property

Public Property Let UserCanceld(ByVal vNewValue As Boolean)
    m_UserCanceld = vNewValue
End Property

Private Sub ISBXPBtnOK_Click()

    With Me.FG

        If .Col = -1 Then Exit Sub
        If .Row = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ItemID"))) = 0 Then
            Exit Sub
        End If

        LngChooseItemID = val(.TextMatrix(.Row, .ColIndex("ItemID")))
        Me.UserCanceld = False
        Me.Hide
    End With

End Sub
