VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowCol 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈ŸÂ«— «·√⁄„œ…"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3000
   Icon            =   "FrmShowCol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   3000
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
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2145
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   2835
      _cx             =   5001
      _cy             =   3784
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmShowCol.frx":038A
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
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Top             =   2730
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
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
      ButtonImage     =   "FrmShowCol.frx":044D
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   2
      Top             =   2730
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
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
      ButtonImage     =   "FrmShowCol.frx":07E7
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√⁄„œ…"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "FrmShowCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

     
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()

    Dim BGround As New ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    On Error GoTo ErrTrap
    Set FG.WallPaper = BGround.Picture
    CenterForm Me

    FormPostion Me, GetPostion
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnCancel_Click()
    FrmMainPriceList.GetMeSetting
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    SaveSetting StrAppRegPath, "ShowCol", "ShowItemID", FG.TextMatrix(0, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowItemCode", FG.TextMatrix(1, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowQty", FG.TextMatrix(2, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowDefalutPrice", FG.TextMatrix(3, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowLastUpdate", FG.TextMatrix(4, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowCustomerPrice", FG.TextMatrix(5, FG.ColIndex("show"))
    SaveSetting StrAppRegPath, "ShowCol", "ShowDealerPrice", FG.TextMatrix(6, FG.ColIndex("show"))

    Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Show & Hide Columns"
    Lbl(0).Caption = "Columns"
    XPBtnOK.Caption = "&Ok"
    XPBtnCancel.Caption = "&Cancel"

    With Me.FG
        FG.TextMatrix(0, 1) = "Show Item ID"
        FG.TextMatrix(1, 1) = "Show Item Code"
        FG.TextMatrix(2, 1) = "Show Item Quantity"
        FG.TextMatrix(3, 1) = "Show Saling Price"
        FG.TextMatrix(4, 1) = "Show LastUpdate"
        FG.TextMatrix(5, 1) = "Show Customer Price"
        FG.TextMatrix(6, 1) = "Show Dealer Price"
        .AutoSize 0, .Cols - 1, False
    End With

End Sub
