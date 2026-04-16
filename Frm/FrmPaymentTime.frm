VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPaymentTime 
   Caption         =   " ‰»ÌÂ«  «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   HelpContextID   =   420
   Icon            =   "FrmPaymentTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7500
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9735
      _cx             =   17171
      _cy             =   13229
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      GridRows        =   7
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPaymentTime.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton CmdOptions 
         Height          =   390
         Left            =   8310
         TabIndex        =   12
         Top             =   30
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ŒÌ«—« ..."
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
         ButtonImage     =   "FrmPaymentTime.frx":0439
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         RightToLeft     =   -1  'True
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   765
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   30
         Width           =   765
         Begin VB.CommandButton Cmd_Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2E9E9&
            Height          =   675
            Left            =   960
            MaskColor       =   &H00FF0000&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.Image Image1 
            Height          =   735
            Left            =   0
            Picture         =   "FrmPaymentTime.frx":07D3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   25
         Left            =   6660
         Top             =   0
      End
      Begin ImpulseButton.ISButton CmdRef 
         Height          =   390
         Left            =   810
         TabIndex        =   9
         Top             =   30
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ «·»Ì«‰« "
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
         ButtonImage     =   "FrmPaymentTime.frx":164F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   375
         Index           =   1
         Left            =   810
         TabIndex        =   8
         Top             =   3810
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
         BackColor       =   14871017
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPaymentTime.frx":19E9
         ColorButton     =   14871017
         ColorHoverText  =   0
         DrawFocusRectangle=   0   'False
         ColorToggledText=   0
         ColorToggledHoverText=   0
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   360
         Index           =   0
         Left            =   810
         TabIndex        =   7
         Top             =   435
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
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
         ButtonImage     =   "FrmPaymentTime.frx":1D83
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7020
         Width           =   9675
         _cx             =   17066
         _cy             =   794
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
         BorderWidth     =   2
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
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   60
            Width           =   3885
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   390
            Left            =   195
            TabIndex        =   3
            Top             =   15
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   688
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
            ButtonImage     =   "FrmPaymentTime.frx":211D
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
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg1 
         Height          =   2985
         Left            =   30
         TabIndex        =   4
         Top             =   810
         Width           =   9675
         _cx             =   17066
         _cy             =   5265
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPaymentTime.frx":24B7
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
      Begin VSFlex8UCtl.VSFlexGrid Fg2 
         Height          =   2805
         Left            =   30
         TabIndex        =   15
         Top             =   4200
         Width           =   9675
         _cx             =   17066
         _cy             =   4948
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
         BackColorBkg    =   16777215
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPaymentTime.frx":279A
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√Ê—«ﬁ «·„«·Ì…:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Index           =   1
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3810
         Width           =   2625
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·√Ê—«ﬁ «·„«·Ì…:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Index           =   0
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   435
         Width           =   2625
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„»«·€ Õ«‰ Êﬁ  ”œ«œÂ«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   1
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   3810
         Width           =   3585
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„»«·€ Õ«‰ Êﬁ  «” ·«„Â«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   435
         Width           =   3585
      End
   End
End
Attribute VB_Name = "FrmPaymentTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrAlramSoundPath As String
Dim BolPlaySound As Boolean
Dim SngTimer As Single

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOptions_Click()
  '  FrmAlramOptions.show vbModal
End Sub

Private Sub CmdPrint_Click(Index As Integer)

    If DoPremis(Do_Print, Me.name, True) = False Then
        Exit Sub
    End If
        
    Dim cNoteReport As ClsNotesReports

    If Index = 0 Then
        Set cNoteReport = New ClsNotesReports
        cNoteReport.ShowCompanyDebitValues 1, Null, Date, False, WindowTarget
        Set cNoteReport = Nothing
    ElseIf Index = 1 Then
        Set cNoteReport = New ClsNotesReports
        cNoteReport.ShowCompanyDebitValues 0, Null, Date, False, WindowTarget
        Set cNoteReport = Nothing
    End If

End Sub

Private Sub CmdRef_Click()
    LoadData
End Sub

Private Sub Fg1_DblClick()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long

    With Me.FG1
    
        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Col = .ColIndex("CusName") Then
            If .TextMatrix(.Row, .ColIndex("CusID")) = "" Then Exit Sub
            LngCusID = val(.TextMatrix(.Row, .ColIndex("CusID")))
            ShowCusBalDailog LngCusID, 0
        ElseIf .Col = .ColIndex("TransactionTypeName") Or .Col = .ColIndex("Transaction_Serial") Then

            If val(.TextMatrix(.Row, .ColIndex("Transaction_Type"))) = 1 Then
                Load FrmBillBuy
                FrmBillBuy.Retrive val(.TextMatrix(.Row, .ColIndex("TransactionsID")))
                FrmBillBuy.show
            ElseIf val(.TextMatrix(.Row, .ColIndex("Transaction_Type"))) = 2 Then
                Load frmsalebill
                frmsalebill.Retrive val(.TextMatrix(.Row, .ColIndex("TransactionsID")))
                frmsalebill.show
            End If
        End If

    End With

End Sub

Private Sub fg2_Click()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long

    With Me.fg2
    
        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Col = .ColIndex("CusName") Then
            If .TextMatrix(.Row, .ColIndex("CusID")) = "" Then Exit Sub
        
            LngCusID = val(.TextMatrix(.Row, .ColIndex("CusID")))
            ShowCusBalDailog LngCusID, 0
        ElseIf .Col = .ColIndex("TransactionTypeName") Or .Col = .ColIndex("Transaction_Serial") Then

            If val(.TextMatrix(.Row, .ColIndex("Transaction_Type"))) = 1 Then
                Load FrmBillBuy
                FrmBillBuy.Retrive val(.TextMatrix(.Row, .ColIndex("TransactionsID")))
                FrmBillBuy.show
            ElseIf val(.TextMatrix(.Row, .ColIndex("Transaction_Type"))) = 2 Then
                Load frmsalebill
                frmsalebill.Retrive val(.TextMatrix(.Row, .ColIndex("TransactionsID")))
                frmsalebill.show
            End If
        End If

    End With

End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap

    If BolPlaySound = True Then
        '    If Dir(StrAlramSoundPath) <> "" Then
        '        PlaySoundEx StrAlramSoundPath, False, True
        '    End If
    End If

    ShowDynamicHelp Me.HelpContextID
    Exit Sub
ErrTrap:
    'App.Path & "\Sound\ALARM3.WAV"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If CmdExit.Enabled = False Then Exit Sub
            CmdExit_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim BGround As New ClsBackGroundPic
    Dim rs As ADODB.Recordset

    CenterForm Me

    FormPostion Me, GetPostion
    Set FG1.WallPaper = BGround.Picture
    Set fg2.WallPaper = BGround.Picture
    FG1.Rows = FG1.FixedRows
    fg2.Rows = fg2.FixedRows
    'Me.ChartPay.PointLabels = False
    'Me.ChartRecv.PointLabels = False
    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Private Sub ChangeLang()
    CmdOptions.Visible = False

    Me.Caption = "Securities outstanding"
    CmdOptions.Caption = "Option"
    CmdRef.Caption = "Reresh"
    LblCaption(0).Caption = "Amounts To receipt"
    lbl(0).Caption = "No. of securities"
    CmdPrint(0).Caption = "Print"
    LblCaption(1).Caption = "Amounts To Payed"
    lbl(1).Caption = "No. of securities"
    CmdPrint(1).Caption = "Print"
    ChkShow.Caption = "Dont Show at start"
    CmdExit.Caption = "Exit"

    With Me.FG1
        .TextMatrix(0, .ColIndex("RequiredValue")) = "Required Value"
        .TextMatrix(0, .ColIndex("Note_Value")) = "Note Value"
        .TextMatrix(0, .ColIndex("LateInterval")) = "Late Interval"
        .TextMatrix(0, .ColIndex("PreRelease")) = "PreRelease"
        .TextMatrix(0, .ColIndex("CusName")) = "Cust. Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type  "
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction Date"
        .TextMatrix(0, .ColIndex("NotesTypeName")) = "Notes Type"
    End With
 
    With Me.fg2
        .TextMatrix(0, .ColIndex("RequiredValue")) = "Required Value"
        .TextMatrix(0, .ColIndex("Note_Value")) = "Note Value"
        .TextMatrix(0, .ColIndex("LateInterval")) = "Late Interval"
        .TextMatrix(0, .ColIndex("PreRelease")) = "PreRelease"
        .TextMatrix(0, .ColIndex("CusName")) = "Cust. Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type  "
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction Date"
        .TextMatrix(0, .ColIndex("NotesTypeName")) = "Notes Type"
    End With

End Sub

Private Sub Form_Resize()
    'If Me.WindowState = vbMaximized Then
    '    Me.ChartRecv.LegendBox = True
    '    Me.ChartPay.LegendBox = True
    'Else
    '    Me.ChartRecv.LegendBox = False
    '    Me.ChartPay.LegendBox = False
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion

    If ChkShow.value = vbChecked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowPayment", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowPayment", True
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LblCaption_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 0 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If

        Exit Sub
    End If

ErrTrap:
End Sub

Private Sub LoadData(Optional BolShowMsg As Boolean = False)
    ShowCurrencyAlarm True
End Sub

Public Sub ApplySetting()
    Dim IntDateDiff As Integer
    Dim rs As ADODB.Recordset
    Dim RowNum As Integer
    Dim SngColor1 As Single, SngColor2 As Single, SngColor3 As Single

    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If rs("PlayNotesAlramSound").value = 1 Then
        BolPlaySound = True
    Else
        BolPlaySound = False
    End If

    StrAlramSoundPath = IIf(IsNull(rs("AlramSoundFilePath").value), App.path & "\Sound\ALARM3.WAV", rs("AlramSoundFilePath").value)

    If rs("EnableNotesAlramColors").value = 1 Then
        SngColor1 = IIf(IsNull(rs("Color1").value), vbWhite, rs("Color1").value)
        SngColor2 = IIf(IsNull(rs("Color2").value), vbWhite, rs("Color2").value)
        SngColor3 = IIf(IsNull(rs("Color3").value), vbWhite, rs("Color3").value)
    Else
        Me.FG1.Clear flexClearScrollable, flexClearFormatting
        Me.fg2.Clear flexClearScrollable, flexClearFormatting
        Exit Sub
    End If

    With Me.FG1

        For RowNum = .FixedRows To .Rows - 1
            IntDateDiff = DateDiff("d", .TextMatrix(RowNum, .ColIndex("DueDate")), Date)

            If IntDateDiff > 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor1
            ElseIf IntDateDiff = 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor2
            Else
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor3
            End If

        Next RowNum

    End With

    With Me.fg2

        For RowNum = .FixedRows To .Rows - 1
            IntDateDiff = DateDiff("d", .TextMatrix(RowNum, .ColIndex("DueDate")), Date)

            If IntDateDiff > 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor1
            ElseIf IntDateDiff = 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor2
            Else
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor3
            End If

        Next RowNum

    End With

End Sub
