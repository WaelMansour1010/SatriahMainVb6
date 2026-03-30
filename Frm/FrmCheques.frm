VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChecks 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ›«’Ì· «·‘Ìﬂ« "
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4545
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   855
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8085
      _cx             =   14261
      _cy             =   1508
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1500
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   420
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   0
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox TxtRowNumber 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
         Width           =   375
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Top             =   420
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         ButtonStyle     =   1
         Caption         =   "≈÷«›…"
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
      Begin VB.TextBox TxtCheckValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5010
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   2205
      End
      Begin MSDataListLib.DataCombo DCboBankName 
         Height          =   315
         Left            =   5010
         TabIndex        =   3
         Top             =   450
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox TxtCheckNumber 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DtpDueDate 
         Height          =   345
         Left            =   2040
         TabIndex        =   4
         Top             =   450
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   609
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·‘Ìﬂ"
         Height          =   255
         Index           =   3
         Left            =   7230
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
         Height          =   405
         Index           =   2
         Left            =   4260
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   390
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   315
         Index           =   1
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   255
         Index           =   0
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid FgCheques 
      Height          =   3045
      Left            =   30
      TabIndex        =   6
      Top             =   900
      Width           =   8085
      _cx             =   14261
      _cy             =   5371
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCheques.frx":0000
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1140
      TabIndex        =   7
      Top             =   4080
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonStyle     =   1
      Caption         =   "Õ›Ÿ"
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
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   4080
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   225
      Index           =   7
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
      Height          =   225
      Index           =   6
      Left            =   4170
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4140
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   225
      Index           =   5
      Left            =   5790
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4140
      Width           =   1305
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·‘Ìﬂ« "
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
      Height          =   225
      Index           =   4
      Left            =   7140
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4140
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   8160
      Y1              =   4020
      Y2              =   4020
   End
End
Attribute VB_Name = "FrmChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PutFg As VSFlex8UCtl.vsFlexGrid

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            AddNewRow

        Case 1
            SaveData

        Case 2
            Unload Me
    End Select

End Sub

Private Sub FgCheques_DblClick()

    With Me.FgCheques

        If .Row <= 0 Then Exit Sub
        If .Col <= 0 Then Exit Sub
        Me.TxtRowNumber.text = .Row
        TxtCheckValue.text = .TextMatrix(.Row, .ColIndex("CheckValue"))
        TxtCheckNumber.text = .TextMatrix(.Row, .ColIndex("CheckNumber"))
        Me.DcboBankName.BoundText = .TextMatrix(.Row, .ColIndex("BankID"))
        Me.DTPDueDate.value = .TextMatrix(.Row, .ColIndex("DueDate"))
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As ClsBackGroundPic
    '----------------------------------------------
    Set Me.Icon = mdifrmmain.ImgLstTree.ListImages("Currency").ExtractIcon
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Plus").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    Cmd(0).ButtonPositionImage = impRightOfText
    Cmd(1).ButtonPositionImage = impRightOfText
    Cmd(2).ButtonPositionImage = impRightOfText
    '----------------------------------------------
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBanks Me.DcboBankName
    Set GrdBack = New ClsBackGroundPic
    SetDtpickerDate Me.DTPDueDate

    With Me.FgCheques
        .RowHeightMin = 300
        .Rows = .FixedRows
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    CenterForm Me

    FormPostion Me, GetPostion
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)

    Select Case Index

        Case 5, 7

            If val(lbl(Index).Caption) > 0 Then
                Me.lbl(Index).ToolTipText = WriteNo(Me.lbl(Index).Caption, 0)
            End If

    End Select

End Sub

Private Sub TxtCheckValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCheckValue.text, 0)
End Sub

Private Sub AddNewRow()
    Dim Msg As String
    Dim LngNewFgRow As Long
    Dim LngFindRow As Long
    Dim IntRes As Integer

    If val(Me.TxtCheckValue.text) = 0 Then
        Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·‘Ìﬂ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.TxtCheckValue.SetFocus
        Exit Sub
    End If

    If Trim$(Me.TxtCheckNumber.text) = "" Then
        Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.TxtCheckNumber.SetFocus
        Exit Sub
    End If

    If val(Me.DcboBankName.BoundText) = 0 Then
        Msg = "ÌÃ» ≈Œ Ì«— «”„ «·»‰ﬂ...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DcboBankName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.DTPDueDate.value = Date Then
        Msg = " «—ÌŒ «·√” Õﬁ«ﬁ €Ì— „ﬁ»Ê· (·«Ì„ﬂ‰ «‰ Ì”«ÊÏ  «—ÌŒ «·ÌÊ„)..!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.DTPDueDate.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Check_CheckNum(Me.TxtCheckNumber.text, val(Me.XPTxtBillID.text), Me.TxtModFlg.text, 0) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If val(Me.TxtRowNumber.text) = 0 Then
        LngFindRow = Me.FgCheques.FindRow(Trim$(Me.TxtCheckNumber.text), Me.FgCheques.FixedRows, FgCheques.ColIndex("CheckNumber"), False, True)

        If LngFindRow <> -1 Then
            Msg = "·«Ì„ﬂ‰  ﬂ—«— ‰›” —ﬁ„ «·‘Ìﬂ ...!!!"
            Msg = Msg & Chr(13) & "Â·  —Ìœ «·≈” »œ«· ...øøø"
            IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

            If IntRes = vbYes Then
                LngNewFgRow = LngFindRow
            Else
                Exit Sub
            End If

        Else
            LngNewFgRow = ModFgLib.SetFgForNewRow(FgCheques, FgCheques.ColIndex("CheckNumber"))
        End If

    Else
        LngNewFgRow = val(Me.TxtRowNumber.text)
    End If

    With Me.FgCheques
        .TextMatrix(LngNewFgRow, .ColIndex("CheckValue")) = val(Me.TxtCheckValue.text)
        .TextMatrix(LngNewFgRow, .ColIndex("CheckNumber")) = Trim$(Me.TxtCheckNumber.text)
        .TextMatrix(LngNewFgRow, .ColIndex("BankID")) = Me.DcboBankName.BoundText
        .TextMatrix(LngNewFgRow, .ColIndex("BankName")) = Me.DcboBankName.text
        .TextMatrix(LngNewFgRow, .ColIndex("DueDate")) = DisplayDate(Me.DTPDueDate.value)
        .AutoSize 0, .Cols - 1, False

        If FgCheques.Rows > 1 Then
            Me.lbl(5).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(7).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(5).Caption = 0
            Me.lbl(7).Caption = 0
        End If

        ModFgLib.ReSerialGrid Me.FgCheques, Me.FgCheques.ColIndex("Serial")
    End With

    Me.TxtRowNumber.text = 0
    Me.TxtCheckValue.SetFocus
End Sub

Private Sub SaveData()
    Dim i As Long
    Dim Msg As String

    If Me.FgCheques.Rows <= 1 Then
        Msg = "ÌÃ» ≈œŒ«· ‘Ìﬂ Ê«Õœ ⁄·Ï «·√ﬁ·...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If ModFgLib.ItemsInGrid(Me.FgCheques, Me.FgCheques.ColIndex("CheckValue")) = -1 Then
        Msg = "ÌÃ» ≈œŒ«· ‘Ìﬂ Ê«Õœ ⁄·Ï «·√ﬁ·...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Not PutFg Is Nothing Then
        PutFg.Clear flexClearScrollable, flexClearEverything
        PutFg.Rows = Me.FgCheques.Rows

        For i = Me.FgCheques.FixedRows To Me.FgCheques.Rows - 1
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("CheckValue")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("CheckValue"))
            
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("CheckNumber")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("CheckNumber"))
            
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("BankID")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("BankID"))
            
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("BankName")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("BankName"))
            
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("DueDate")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("DueDate"))
            
            Me.PutFg.TextMatrix(i, Me.PutFg.ColIndex("ReleaseDate")) = Me.FgCheques.TextMatrix(i, Me.FgCheques.ColIndex("ReleaseDate"))
            
        Next i

        PutFg.AutoSize 0, PutFg.Cols - 1, False
        Unload Me
    End If

End Sub
