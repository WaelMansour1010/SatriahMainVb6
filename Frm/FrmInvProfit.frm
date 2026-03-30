VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmInvProfit 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’«›Ï —»Õ ›« Ê—…"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "FrmInvProfit.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDisplayType 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   570
      Width           =   3255
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷  ›’Ì·Ì"
         Height          =   315
         Index           =   1
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1425
      End
      Begin VB.OptionButton OptType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ „Œ ’—"
         Height          =   315
         Index           =   0
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   555
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8085
      _cx             =   14261
      _cy             =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "’«›Ï —»Õ ›« Ê—…"
      Align           =   0
      AutoSizeChildren=   0
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
   End
   Begin VSFlex8UCtl.VSFlexGrid FgItems 
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   1260
      Width           =   8085
      _cx             =   14261
      _cy             =   5424
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmInvProfit.frx":058A
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
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   5490
      Width           =   795
      _ExtentX        =   1402
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
      BackStyle       =   0
      ButtonImage     =   "FrmInvProfit.frx":07C8
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
   Begin VB.Label lBL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "’«›Ï —»Õ «·›« Ê—…"
      Height          =   345
      Index           =   2
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4380
      Width           =   1365
   End
   Begin VB.Label lblInvProf 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4380
      Width           =   1545
   End
   Begin VB.Label lBL 
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
      ForeColor       =   &H00000080&
      Height          =   435
      Index           =   1
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   780
      Width           =   2925
   End
   Begin VB.Label lBL 
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
      ForeColor       =   &H00000080&
      Height          =   1125
      Index           =   0
      Left            =   870
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4770
      Width           =   7155
   End
End
Attribute VB_Name = "FrmInvProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CostMethod As StockCostType

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub FgItems_DblClick()

    If Me.CostMethod = ModernWeightAverage Then

        With Me.FgItems

            If val(.TextMatrix(.Row, .ColIndex("ItemID"))) <> 0 Then
                Load FrmItemCostShow
                FrmItemCostShow.DcboItemName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
                FrmItemCostShow.DoAction
                FrmItemCostShow.Show
                FrmItemCostShow.ZOrder 0
            End If

        End With

    End If

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Dim GrdBack As ClsBackGroundPic
    CenterForm Me
    Set GrdBack = New ClsBackGroundPic

    With Me.FgItems
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & Chr(13) & "»«·‰”»… ··√’‰«› «· Ï ·«   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«· ›«‰ «·»—‰«„Ã ÌﬁÊ„ »Õ”«» ”⁄— «· ﬂ·›… »‰«¡ ⁄·Ï «Œ— ”⁄— ‘—«¡ ”Ã· ·Â–« «·’‰›"
    Msg = Msg & Chr(13) & "Ê»«·‰”»… ··√’‰«› «· Ï   ⁄«„· »‰Ÿ«„ «·”Ì—Ì«· ›«‰ «·»—‰«„Ã Ì” œ· ⁄·Ï ”⁄— «·’‰› »œ·«·… «·”Ì—Ì«· «·Œ«’ »Â"
    'Msg = Msg & Chr(13) & ""
    Me.lbl(0).Caption = Msg
    Msg = "«·⁄—÷ «· ›’Ì·Ì ÌﬁÊ„ »⁄—÷ ‰Ê⁄ «·⁄„·Ì… Ê—ﬁ„ «·„”·”· «·Œ«’ »Â«"
    Me.lbl(1).Caption = Msg

    Me.OptType(1).value = True
    OptType_Click (1)
End Sub

Private Sub OptType_Click(Index As Integer)

    With Me.FgItems
        .ColHidden(.ColIndex("TransType")) = OptType(0).value
        .ColHidden(.ColIndex("TransSerial")) = OptType(0).value
        '.ColHidden(.ColIndex("")) = OptType(0).Value
    End With

End Sub

Public Property Get CostMethod() As StockCostType
    CostMethod = m_CostMethod
End Property

Public Property Let CostMethod(ByVal vNewValue As StockCostType)
    m_CostMethod = vNewValue

    If m_CostMethod = ModernWeightAverage Or m_CostMethod = WeightAverage Then
        Me.FraDisplayType.Visible = False
        Me.lbl(1).Visible = False
    
        With Me.FgItems
            .ColHidden(.ColIndex("TransType")) = True
            .ColHidden(.ColIndex("TransSerial")) = True
        End With

    End If

End Property
