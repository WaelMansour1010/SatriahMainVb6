VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowTransNotes 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·⁄„·Ì«  «·„«·Ì…"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "FrmShowTransNotes.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   7335
      _cx             =   12938
      _cy             =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ForeColor       =   128
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " ›«’Ì· «·Õ—ﬂ…"
      Align           =   0
      AutoSizeChildren=   0
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
      Begin ImpulseAniLabel.ISAniLabel LblLink 
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   450
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   450
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmShowTransNotes.frx":038A
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   12
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   450
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   255
         Index           =   9
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   6450
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   450
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   7
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   6450
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   210
         Width           =   825
      End
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   4860
      Width           =   825
      _ExtentX        =   1455
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
      ButtonImage     =   "FrmShowTransNotes.frx":04EC
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   3075
      Left            =   0
      TabIndex        =   4
      Top             =   1290
      Width           =   7335
      _cx             =   12938
      _cy             =   5424
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   12640511
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
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmShowTransNotes.frx":0886
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
      PictureType     =   1
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ:"
      Height          =   285
      Index           =   4
      Left            =   2460
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4410
      Width           =   405
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   5
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4410
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï :"
      Height          =   285
      Index           =   3
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4410
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   2
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4410
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "»Ì«‰«  «·⁄„·Ì«  «·„«·Ì… «·ﬁ«∆„… ⁄·Ï «·Õ—ﬂ…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   450
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6870
      Picture         =   "FrmShowTransNotes.frx":0A08
      Top             =   4830
      Width           =   480
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4950
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   7350
      Y1              =   4770
      Y2              =   4770
   End
End
Attribute VB_Name = "FrmShowTransNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IntMode As Integer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Fg_DblClick()
    Dim LngNoteID As Long
    Dim LngNoteType As Long
    Dim LngTransID As Long
    Dim LngTrasType As Long

    With Me.Fg

        If .Row <= 0 Then Exit Sub
        If .Col <= 0 Then Exit Sub
        If Me.IntMode = 0 Then
            LngNoteID = val(.TextMatrix(.Row, .ColIndex("NoteID")))

            If LngNoteID <> 0 Then
                LngNoteType = val(.TextMatrix(.Row, .ColIndex("NoteType")))

                If LngNoteType = 0 Then Exit Sub

                Select Case LngNoteType

                    Case 4
                        Load FrmCashing
                        FrmCashing.Retrive LngNoteID
                        FrmCashing.ZOrder 0

                    Case 5
                        Load FrmPayments
                        FrmPayments.Retrive LngNoteID
                        FrmPayments.ZOrder 0

                    Case 9, 10
                        Load FrmDiscounts
                        FrmDiscounts.Retrive LngNoteID
                        FrmDiscounts.ZOrder 0
                End Select

            End If

        ElseIf Me.IntMode = 1 Then
            LngTransID = val(.TextMatrix(.Row, .ColIndex("Transaction_ID")))

            If LngTransID <> 0 Then
                LngTrasType = val(.TextMatrix(.Row, .ColIndex("Transaction_Type")))

                If LngTrasType = 0 Then Exit Sub

                Select Case LngTrasType

                    Case 9
                        Load FrmReturnSalling
                        FrmReturnSalling.Retrive LngTransID
                        FrmReturnSalling.ZOrder 0
                End Select

            End If
        End If

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As New ClsBackGroundPic
    Dim Msg As String

    If Me.MDIChild = True Then
        Resize_Form Me
    Else
        CenterForm Me
    End If

    With Me.Fg
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    Msg = "Ì„ﬂ‰ﬂ „‘«Âœ… «Ï Õ—ﬂ… »«· ›’Ì· »«·÷€Ÿ „— Ì‰ „  «·Ì Ì‰ ⁄·ÌÂ«"
    Me.lbl(1).Caption = Msg

    FormPostion Me, GetPostion
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    Select Case Index

        Case 2, 5
            lbl(Index).ToolTipText = WriteNo(Me.lbl(Index).Caption, 0)
    End Select

End Sub

Private Sub LblLink_Click()

    If val(Me.LblLink.Tag) <> 0 Then
        ShowCusBalDailog val(Me.LblLink.Tag), 0
    Else
        Exit Sub
    End If

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    Dim StrTemp  As String

    If val(Me.LblLink.Tag) <> 0 Then
        StrTemp = "≈÷€ÿ Â‰« Õ Ï  Õ’· ⁄·Ï  ﬁ«—Ì— Œ«’… »‹ :" & Trim$(Me.LblLink.Caption)
        Me.LblLink.ToolTipText = StrTemp
    Else
        Me.LblLink.ToolTipText = ""
    End If

End Sub
