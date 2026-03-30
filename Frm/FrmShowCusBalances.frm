VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowCusBalances 
   Caption         =   "⁄—÷ √—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   HelpContextID   =   1001
   Icon            =   "FrmShowCusBalances.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11535
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8265
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11535
      _cx             =   20346
      _cy             =   14579
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
      GridRows        =   4
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmShowCusBalances.frx":058A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   885
         Index           =   0
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   30
         Width           =   11475
         _cx             =   20241
         _cy             =   1561
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   765
            Index           =   1
            Left            =   30
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   30
            Width           =   4740
            _cx             =   8361
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
            Appearance      =   0
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
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Œ›«¡ «·√—’œ… «·’›—Ì…"
               Height          =   315
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   390
               Width           =   3435
            End
            Begin VB.ComboBox CboDisplayType 
               Height          =   315
               Left            =   1260
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   30
               Width           =   3435
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   360
               Index           =   2
               Left            =   30
               TabIndex        =   8
               Top             =   0
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   635
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   " ÕœÌÀ"
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
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   360
               Index           =   1
               Left            =   30
               TabIndex        =   9
               Top             =   390
               Width           =   1215
               _ExtentX        =   2143
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
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ √—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰"
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
            Height          =   420
            Index           =   0
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   30
            Width           =   6630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì„ﬂ‰ﬂ ⁄—÷  ﬁ—Ì— »ﬂ‘› Õ”«» «Ï ⁄„Ì· »«·÷€ÿ ⁄·Ï «”„ «·⁄„Ì· «Ê «·„Ê—œ „— Ì‰ „  «·Ì Ì‰"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Index           =   2
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   465
            Width           =   6630
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   420
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   7815
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
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
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6870
         Left            =   30
         TabIndex        =   1
         Top             =   930
         Width           =   11475
         _cx             =   20241
         _cy             =   12118
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
         FloodColor      =   16744576
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
         Rows            =   50
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmShowCusBalances.frx":0615
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
         AllowUserFreezing=   3
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Index           =   1
         Left            =   1185
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   7815
         Width           =   10320
      End
   End
End
Attribute VB_Name = "FrmShowCusBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub Cmd_Click(Index As Integer)
    Dim cReport As ClsCustemerReport

    Select Case Index

        Case 0
            Unload Me

        Case 1
            Set cReport = New ClsCustemerReport
            cReport.ShowCustsBalances Me.CboDisplayType.ListIndex, WindowTarget
            Set cReport = Nothing

        Case 2
            LoadData
    End Select

End Sub

Private Sub EleMain_DblClick()
    Me.WindowState = IIf(Me.WindowState = vbMaximized, vbNormal, vbMaximized)
End Sub

Private Sub Fg_DblClick()
    Dim LngCusID As Long

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If .Col <> .ColIndex("CUsName") Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("CusID")) = "" Then Exit Sub
        LngCusID = val(.TextMatrix(.Row, .ColIndex("CusID")))
        ShowCusBalDailog LngCusID, 0
    End With

End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       x As Single, _
                       Y As Single)
    Dim LngRow As Long
    Dim LngCusID As Long

    If Button = vbRightButton Then

        With FG
            LngRow = .MouseRow

            If LngRow <= 0 Then Exit Sub
            LngCusID = val(.TextMatrix(LngRow, .ColIndex("CusID")))

            If LngCusID <> 0 Then
                mdifrmmain.MnuCusTools.Tag = LngCusID
                Me.PopupMenu mdifrmmain.MnuCusTools
            End If

        End With

    End If

End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub Form_Load()
    Dim cGrdBack As ClsBackGroundPic
    Set cGrdBack = New ClsBackGroundPic
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Refresh").Picture

    With Me.CboDisplayType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem "«·⁄„·«¡ Ê«·„Ê—œÌ‰"
            .AddItem "√—’œ… «·⁄„·«¡ ›ﬁÿ"
            .AddItem "√—’œ… «·„Ê—œÌ‰ ›ﬁÿ"
            .AddItem "√—’œ… «·„ ⁄·ﬁ« "
            .AddItem "√—’œ… «·⁄„·«¡ Ê«·„Ê—œÌ‰ Ê«·„ ⁄·ﬁ« "
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Customers and Suppliers"
            .AddItem "Customers Only"
            .AddItem "Suppliers Only"
            .AddItem "Other Acccounts"
            .AddItem "All(Customers,Suppliers,Other Acccounts)"
        End If

        .ListIndex = 0
    End With

    With Me.FG
        .MergeCells = flexMergeFixedOnly

        If SystemOptions.UserInterface = ArabicInterface Then
            .Cell(flexcpText, 0, .ColIndex("Serial"), 1, .ColIndex("Serial")) = "„”·”·"
            .MergeCol(.ColIndex("Serial")) = True
        
            .Cell(flexcpText, 0, .ColIndex("CusID"), 1, .ColIndex("CusID")) = "ﬂÊœ «·⁄„Ì· «Ê «·„Ê—œ"
            .MergeCol(.ColIndex("CusID")) = True
            .Cell(flexcpText, 0, .ColIndex("CusName"), 1, .ColIndex("CusName")) = "«·√”„"
            .MergeCol(.ColIndex("CusName")) = True
            .Cell(flexcpText, 0, .ColIndex("Type"), 1, .ColIndex("Type")) = "‰Ê⁄Â"
            .MergeCol(.ColIndex("Type")) = True
        
            .Cell(flexcpText, 0, .ColIndex("OpenDebit"), 0, .ColIndex("OpenCredit")) = "«·—’Ìœ «·√›  «ÕÏ"
            .AutoSize 0, .Cols - 1, False
            .MergeRow(0) = True
            .Cell(flexcpText, 1, .ColIndex("OpenDebit"), 1, .ColIndex("OpenDebit")) = "„œÌ‰"
            .Cell(flexcpText, 1, .ColIndex("OpenCredit"), 1, .ColIndex("OpenCredit")) = "œ«∆‰"
        
            .Cell(flexcpText, 0, .ColIndex("CurrentDebit"), 0, .ColIndex("CurrentCredit")) = "«·—’Ìœ «·Õ«·Ï"
            .MergeRow(0) = True
            .Cell(flexcpText, 1, .ColIndex("CurrentDebit"), 1, .ColIndex("CurrentDebit")) = "„œÌ‰"
            .Cell(flexcpText, 1, .ColIndex("CurrentCredit"), 1, .ColIndex("CurrentCredit")) = "œ«∆‰"
        
            .Cell(flexcpText, 0, .ColIndex("OpenBalanceDate"), 1, .ColIndex("OpenBalanceDate")) = " «—ÌŒ  ”ÃÌ· «·—’Ìœ «·√›  «ÕÏ"
            .MergeCol(.ColIndex("OpenBalanceDate")) = True
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .Cell(flexcpText, 0, .ColIndex("Serial"), 1, .ColIndex("Serial")) = "Serial"
            .MergeCol(.ColIndex("Serial")) = True
        
            .Cell(flexcpText, 0, .ColIndex("CusID"), 1, .ColIndex("CusID")) = "Code"
            .MergeCol(.ColIndex("CusID")) = True
            .Cell(flexcpText, 0, .ColIndex("CusName"), 1, .ColIndex("CusName")) = "Dlear Name"
            .MergeCol(.ColIndex("CusName")) = True
            .Cell(flexcpText, 0, .ColIndex("Type"), 1, .ColIndex("Type")) = "Dlear Type"
            .MergeCol(.ColIndex("Type")) = True
        
            .Cell(flexcpText, 0, .ColIndex("OpenDebit"), 0, .ColIndex("OpenCredit")) = "Opening Balance"
            .AutoSize 0, .Cols - 1, False
            .MergeRow(0) = True
            .Cell(flexcpText, 1, .ColIndex("OpenDebit"), 1, .ColIndex("OpenDebit")) = "Debit"
            .Cell(flexcpText, 1, .ColIndex("OpenCredit"), 1, .ColIndex("OpenCredit")) = "Credit"
        
            .Cell(flexcpText, 0, .ColIndex("CurrentDebit"), 0, .ColIndex("CurrentCredit")) = "Current Balance"
            .MergeRow(0) = True
            .Cell(flexcpText, 1, .ColIndex("CurrentDebit"), 1, .ColIndex("CurrentDebit")) = "Debit"
            .Cell(flexcpText, 1, .ColIndex("CurrentCredit"), 1, .ColIndex("CurrentCredit")) = "Credit"
            .Cell(flexcpText, 0, .ColIndex("OpenBalanceDate"), 1, .ColIndex("OpenBalanceDate")) = "Opening Balance Date"
            .MergeCol(.ColIndex("OpenBalanceDate")) = True
        End If

        .AutoSize 0, .ColIndex("Type"), False
        .ColWidth(.ColIndex("OpenDebit")) = 1050
        .ColWidth(.ColIndex("OpenCredit")) = 1050
        .ColWidth(.ColIndex("CurrentDebit")) = 1050
        .ColWidth(.ColIndex("CurrentCredit")) = 1050
        .WallPaper = cGrdBack.Picture
    End With

    Chk.value = vbChecked
    LoadData
    Me.Height = 9500
    Me.Width = 11600
    Resize_Form Me
End Sub

Private Sub LoadData()
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim SngCurrentBalance As Single
    Dim SngMaxDebit As Single
    Dim SngMaxCredit As Single
    Dim SngOpenBalance As Single
    Dim BolFrmLoaded As Boolean
    Dim cProgress As ClsProgress
    Dim SngTimer As Single

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblCustemers.CusID, TblCustemers.CusName,TblCustemers.Type," & "TblCustemers.OpenBalance, TblCustemers.OpenBalanceType, TblCustemers.OpenBalanceDate, " & "Sum(QryCustomerBalance.Note_Value * QryCustomerBalance.CreditOrDebit) AS CustomerAccount"
        StrSQL = StrSQL + " FROM QryCustomerBalance RIGHT JOIN TblCustemers ON " & " QryCustomerBalance.CusID = TblCustemers.CusID"

        If Me.CboDisplayType.ListIndex = 0 Then
            StrSQL = StrSQL + " Where (TblCustemers.Type=1 OR TblCustemers.Type=2)"
        ElseIf Me.CboDisplayType.ListIndex = 1 Then
            StrSQL = StrSQL + " Where TblCustemers.Type=1"
        ElseIf Me.CboDisplayType.ListIndex = 2 Then
            StrSQL = StrSQL + " Where TblCustemers.Type=2"
        ElseIf Me.CboDisplayType.ListIndex = 3 Then
            StrSQL = StrSQL + " Where TblCustemers.Type=3"
        Else
            StrSQL = StrSQL + " Where (TblCustemers.Type=1 OR TblCustemers.Type=2 OR TblCustemers.Type=3)"
        End If

        StrSQL = StrSQL + " GROUP BY TblCustemers.CusID, TblCustemers.CusName, TblCustemers.OpenBalance," & "TblCustemers.OpenBalanceType, TblCustemers.OpenBalanceDate, TblCustemers.Type "

        If Me.Chk.value = vbChecked Then
            'StrSQL = StrSQL + " HAVING  Sum(QryCustomerBalance.Note_Value*QryCustomerBalance.CreditOrDebit) <> 0"
        End If

        StrSQL = StrSQL + " Order By TblCustemers.CusID"
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adAsyncFetch
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT TOP 100 PERCENT dbo.TblCustemers.CusID, dbo.TblCustemers.CusName," & "dbo.TblCustemers.Type, dbo.TblCustemers.OpenBalance, " & "dbo.TblCustemers.OpenBalanceType, dbo.TblCustemers.OpenBalanceDate," & "SUM(dbo.QryCustomerBalance.Note_Value * dbo.QryCustomerBalance.CreditOrDebit) As CustomerAccount"
        StrSQL = StrSQL + " FROM dbo.QryCustomerBalance(0) RIGHT JOIN " & "dbo.TblCustemers ON dbo.QryCustomerBalance.CusID = dbo.TblCustemers.CusID "

        If Me.CboDisplayType.ListIndex = 0 Then
            StrSQL = StrSQL + " Where (dbo.TblCustemers.Type=1 OR dbo.TblCustemers.Type=2) "
        ElseIf Me.CboDisplayType.ListIndex = 1 Then
            StrSQL = StrSQL + " Where dbo.TblCustemers.Type=1"
        ElseIf Me.CboDisplayType.ListIndex = 2 Then
            StrSQL = StrSQL + " Where dbo.TblCustemers.Type=2"
        ElseIf Me.CboDisplayType.ListIndex = 3 Then
            StrSQL = StrSQL + " Where dbo.TblCustemers.Type=3"
        End If

        StrSQL = StrSQL + " GROUP BY dbo.TblCustemers.CusID, dbo.TblCustemers.CusName," & "dbo.TblCustemers.OpenBalance , dbo.TblCustemers.OpenBalanceType, " & "dbo.TblCustemers.OpenBalanceDate , dbo.TblCustemers.Type"

        If Me.Chk.value = vbChecked Then
            'StrSQL = StrSQL + " Having (SUM(dbo.QryCustomerBalance.Note_Value * dbo.QryCustomerBalance.CreditOrDebit) <> 0)"
        End If

        StrSQL = StrSQL + " ORDER BY dbo.TblCustemers.CusID"
        rs.CursorLocation = adUseClient
        rs.Properties("Initial Fetch Size") = 2
        rs.Properties("Background Fetch Size") = 4

        DoEvents
        Cn.CommandTimeout = 180
        SngTimer = Timer
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncFetch + adAsyncExecute
        ' adAsyncExecute
        'Rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop
    
    End If

    If BolFrmLoaded = True Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    With Me.FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
    End With

    'MsgBox Timer - SngTimer
    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì‰«‰«  ..."
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.lbl(1).Caption = Msg
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    With Me.FG
        Screen.MousePointer = vbArrowHourglass
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + rs.RecordCount
        rs.MoveFirst
          
        For i = .FixedRows To .Rows - 1
            SngCurrentBalance = 0
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                If rs("Type").value = 1 Then
                    .TextMatrix(i, .ColIndex("Type")) = "⁄„Ì·"
                ElseIf rs("Type").value = 2 Then
                    .TextMatrix(i, .ColIndex("Type")) = "„Ê—œ"
                ElseIf rs("Type").value = 3 Then
                    .TextMatrix(i, .ColIndex("Type")) = "„ ⁄«„·Ê‰"
                End If

            ElseIf SystemOptions.UserInterface = EnglishInterface Then

                If rs("Type").value = 1 Then
                    .TextMatrix(i, .ColIndex("Type")) = "Customer"
                ElseIf rs("Type").value = 2 Then
                    .TextMatrix(i, .ColIndex("Type")) = "Supplier"
                ElseIf rs("Type").value = 3 Then
                    .TextMatrix(i, .ColIndex("Type")) = "Other Accounts"
                End If
            End If

            If Not (IsNull(rs("OpenBalance").value)) Then
                If rs("OpenBalanceType").value = 0 Then '„œÌ‰
                    .TextMatrix(i, .ColIndex("OpenDebit")) = IIf(IsNull(rs("OpenBalance").value), 0, rs("OpenBalance").value)
                    .TextMatrix(i, .ColIndex("OpenCredit")) = 0
                    SngCurrentBalance = -1 * rs("OpenBalance").value
                ElseIf rs("OpenBalanceType").value = 1 Then 'œ«∆‰
                    .TextMatrix(i, .ColIndex("OpenCredit")) = IIf(IsNull(rs("OpenBalance").value), 0, rs("OpenBalance").value)
                    .TextMatrix(i, .ColIndex("OpenDebit")) = 0
                    SngCurrentBalance = rs("OpenBalance").value
                End If

                If Not (IsNull(rs("OpenBalanceDate").value)) Then
                    .TextMatrix(i, .ColIndex("OpenBalanceDate")) = DisplayDate(rs("OpenBalanceDate").value)
                End If
            End If
        
            SngCurrentBalance = IIf(IsNull(rs("CustomerAccount").value), SngCurrentBalance, SngCurrentBalance + rs("CustomerAccount").value)
            SngCurrentBalance = Format(SngCurrentBalance, SystemOptions.SysDefCurrencyForamt)

            If SngCurrentBalance < 0 Then '„œÌ‰
                .TextMatrix(i, .ColIndex("CurrentDebit")) = Abs(SngCurrentBalance)
                .TextMatrix(i, .ColIndex("CurrentCredit")) = 0

                '«·Õ’Ê· ⁄·Ï «⁄·Ï ﬁÌ„… „œÌ‰…
                '·ﬂÏ ‰” Œœ„Â« ›Ï —”„ «·„⁄Ì«— ›Ï «·‹ Grid
                If SngMaxDebit < Abs(SngCurrentBalance) Then
                    SngMaxDebit = Abs(SngCurrentBalance)
                End If

            ElseIf SngCurrentBalance > 0 Then 'œ«∆‰
                .TextMatrix(i, .ColIndex("CurrentCredit")) = SngCurrentBalance
                .TextMatrix(i, .ColIndex("CurrentDebit")) = 0

                If SngMaxCredit < SngCurrentBalance Then
                    SngMaxCredit = Abs(SngCurrentBalance)
                End If

            Else
            End If

            rs.MoveNext
        Next i
    
        If Me.Chk.value = vbChecked Then
            i = .FixedRows

            Do While (i <= .Rows - 1)

                If val(.TextMatrix(i, .ColIndex("CurrentDebit"))) = 0 And val(.TextMatrix(i, .ColIndex("CurrentCredit"))) = 0 Then
                    .RemoveItem i
                    i = .FixedRows
                Else
                    i = i + 1
                End If

            Loop

        End If

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Serial")) = i - 1
        Next i

        .Cell(flexcpFloodColor, .FixedRows, .ColIndex("CurrentDebit"), .Rows - 1, .ColIndex("CurrentDebit")) = &HC0&
        .Cell(flexcpFloodColor, .FixedRows, .ColIndex("CurrentCredit"), .Rows - 1, .ColIndex("CurrentCredit")) = &HFF8080

        If SngMaxDebit <> 0 Then

            For i = .FixedRows To .Rows - 1
                .Cell(flexcpFloodPercent, i, .ColIndex("CurrentDebit")) = 100 * val(.TextMatrix(i, .ColIndex("CurrentDebit"))) / SngMaxDebit
            Next i

        End If

        If SngMaxCredit <> 0 Then

            For i = .FixedRows To .Rows - 1
                .Cell(flexcpFloodPercent, i, .ColIndex("CurrentCredit")) = 100 * val(.TextMatrix(i, .ColIndex("CurrentCredit"))) / SngMaxCredit
            Next i

        End If

        .Refresh
        .AutoSize 0, .ColIndex("Type"), False
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«·⁄œœ:- " & "(" & .Rows - .FixedRows & ")"
            Msg = Msg & " ≈Ã„«·Ï «·√—’œ… «·„œÌ‰…:- " & "(" & .Aggregate(flexSTSum, .FixedRows, .ColIndex("CurrentDebit"), .Rows - 1, .ColIndex("CurrentDebit")) & ")"
            Msg = Msg & " ≈Ã„«·Ï «·√—’œ… «·œ«∆‰…:- " & "(" & .Aggregate(flexSTSum, .FixedRows, .ColIndex("CurrentCredit"), .Rows - 1, .ColIndex("CurrentCredit")) & ")"
        Else
            Msg = "Number:- " & "(" & .Rows - .FixedRows & ")"
            Msg = Msg & "Total Debit Accounts:- " & "(" & .Aggregate(flexSTSum, .FixedRows, .ColIndex("CurrentDebit"), .Rows - 1, .ColIndex("CurrentDebit")) & ")"
            Msg = Msg & " Total Credit Accounts:- " & "(" & .Aggregate(flexSTSum, .FixedRows, .ColIndex("CurrentCredit"), .Rows - 1, .ColIndex("CurrentCredit")) & ")"
        End If

        Me.lbl(1).Caption = Msg
    End With

    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

    Msg = " „ «·√” ⁄·«„ .."
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub ChangeLang()
    Me.Caption = "Customers and Suppliers Balance"
    Me.lbl(0).Caption = Me.Caption
    Me.lbl(2).Caption = "Double click on dealer name to show its  balanca report"
    Chk.Caption = "Hide Zero Balance"
    Cmd(0).Caption = "Exit"
    Cmd(1).Caption = "Print"
    Cmd(2).Caption = "Refresh"
End Sub

Private Sub lbl_Click(Index As Integer)

    Select Case Index

        Case 0, 2
            Me.WindowState = IIf(Me.WindowState = vbMaximized, vbNormal, vbMaximized)
    End Select

End Sub

Private Sub Rs_FetchProgress(ByVal Progress As Long, _
                             ByVal MaxProgress As Long, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)
    Debug.Print Me.name & " Rs_FetchProgress", Progress, MaxProgress; ""
End Sub

