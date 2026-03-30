VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBoxSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì«    «· ÕÊÌ·«  «·„«·Ì…"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "FrmBoxSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTperson 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2160
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3240
      Width           =   2025
   End
   Begin VB.ComboBox CboType 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2490
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3210
      Width           =   1185
   End
   Begin VB.TextBox XPTxtBillNum 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6480
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2490
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2850
      Width           =   2085
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   196608001
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   196608001
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   285
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   4020
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   855
      TabIndex        =   2
      Top             =   4020
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   4020
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   5040
      TabIndex        =   12
      Top             =   2850
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8865
      _cx             =   15637
      _cy             =   4313
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBoxSearch.frx":058A
      ScrollTrack     =   -1  'True
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
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   5280
      TabIndex        =   16
      Top             =   3570
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” ›Ìœ"
      Height          =   315
      Index           =   5
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " — Ì» «·»Ì«‰« :"
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
      Height          =   225
      Index           =   1
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4470
      Width           =   4845
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ :"
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
      Height          =   225
      Index           =   0
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4200
      Width           =   2265
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   2190
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2490
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   2
      Left            =   7860
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3690
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3210
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   315
      Index           =   0
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·≈Ìœ«⁄"
      Height          =   315
      Index           =   3
      Left            =   7950
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2490
      Width           =   765
   End
End
Attribute VB_Name = "FrmBoxSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_SearchNoteType As Integer

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
            GetData

        Case 1
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub

Private Sub Fg_AfterSort(ByVal Col As Long, _
                         Order As Integer)
    Me.lbl(1).Caption = GetFgSortTitle(FG, Col, Order)
    'Me.Lbl(1).Caption = Order
End Sub

Private Sub fg_Click()

    With Me.FG

        If .row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.row, .ColIndex("NoteID"))) = 0 Then
            Exit Sub
        End If

        mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.row, .ColIndex("NoteID")))
    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim cGridBack As New ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion
    Me.DTPFrom.value = Date
    Me.DTPFrom.value = Null
    Me.DTPTo.value = Date
    Me.DTPTo.value = Null

    With Me.FG
        Set .WallPaper = cGridBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.CboType
        .Clear
        .AddItem "“Ì«œ… ›Ï «·‰ﬁœÌ…"
        .AddItem "⁄Ã“ ›Ï «·‰ﬁœÌ…"
        .AddItem "«·ﬂ·"
    End With

    Set Dcombos = New ClsDataCombos
    Dcombos.GetBanks Me.DcboBox
    Dcombos.GetUsers Me.DcboUsers
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    StrSQL = "SELECT Notes.NoteID,Notes.NoteDate , Notes.NoteType, Notes.NoteSerial,Notes.NoteSerial1," & "Notes.Note_Value, Notes.NoteHijriDate,Notes.Remark, TblUsers.UserID, TblUsers.UserName," & "BanksData.BankID, BanksData.BankName BoxName "
    StrSQL = StrSQL + " FROM            BanksData RIGHT OUTER JOIN"
    StrSQL = StrSQL + "                          TblUsers LEFT OUTER JOIN"
   StrSQL = StrSQL + "                       Notes ON TblUsers.UserID = Notes.UserID ON BanksData.BankID = Notes.BankID"
    'StrSQL = StrSQL + " FROM TblUsers INNER JOIN (TblBoxesData   INNER JOIN Notes ON " & "TblBoxesData.BoxID = Notes.BoxID) ON TblUsers.UserID = Notes.UserID "

    BolBegine = True

    If Me.SearchNoteType = 7 Or Me.SearchNoteType = 8 Then
        StrWhere = " Where  Notes.NoteType=" & 14 & ""
    ElseIf Me.SearchNoteType = 11 Then
       ' StrWhere = " Where ( Notes.NoteType= 11 Or  Notes.NoteType= 12)"
        
                If Me.CboType.ListIndex = 0 Then
            StrWhere = " Where  Notes.NoteType=11"
        ElseIf Me.CboType.ListIndex = 1 Then
            StrWhere = " Where  Notes.NoteType=12"
        Else
            StrWhere = " Where  (Notes.NoteType=11 OR Notes.NoteType=12)"
        End If
        
    Else

        If Me.CboType.ListIndex = 0 Then
            StrWhere = " Where  Notes.NoteType=14"
        ElseIf Me.CboType.ListIndex = 1 Then
            StrWhere = " Where  Notes.NoteType=14"
        Else
            StrWhere = " Where  (Notes.NoteType=14 OR Notes.NoteType=14)"
        End If
    End If

    If 1 = 1 Then
        If val(Me.XPTxtBillNum.text) <> 0 Then
            If BolBegine = True Then
                StrWhere = StrWhere & " AND Notes.NoteSerial1=" & (Me.XPTxtBillNum.text) & ""
            Else
                BolBegine = True
                StrWhere = " Where Notes.NoteSerial1=" & (Me.XPTxtBillNum.text) & ""
            End If
        End If

    Else

        If val(Me.XPTxtBillNum.text) <> 0 Then
            If BolBegine = True Then
                StrWhere = StrWhere & " AND Notes.NoteSerial=" & val(Me.XPTxtBillNum.text) & ""
            Else
                BolBegine = True
                StrWhere = " Where Notes.NoteSerial=" & val(Me.XPTxtBillNum.text) & ""
            End If
        End If

    End If

    If Me.DcboBox.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND BanksData.BankID =" & Me.DcboBox.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where BanksData.BankID =" & Me.DcboBox.BoundText & ""
        End If
    End If

    If val(Me.TxtValue.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Notes.Note_Value=" & val(Me.TxtValue.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where Notes.Note_Value=" & val(Me.TxtValue.text) & ""
        End If
    End If

    If Me.DcboUsers.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Notes.UserID=" & Me.DcboUsers.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where Notes.UserID=" & Me.DcboUsers.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DTPFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Notes.NoteDate >=" & SQLDate(Me.DTPFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where Notes.NoteDate >=" & SQLDate(Me.DTPFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Notes.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where Notes.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        End If
    End If
    
    
   If txtperson.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND person LIKE'%" & txtperson.text & "%'"
        Else
            BolBegine = True
            StrWhere = StrWhere & " WHERE person LIKE'%" & txtperson.text & "%'"
        End If
    End If
    
    
    

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & "  AND      Notes.branch_no in(" & Current_branchSql & ")"
    StrSQL = StrSQL & " Order By Notes.NoteID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(0).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    clear_all Me
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(0).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)

                If Not (IsNull(rs("NoteDate").value)) Then
                    .TextMatrix(i, .ColIndex("NoteDate")) = Format(rs("NoteDate").value, "yyyy/M/d")
                End If

                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value), rs("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

End Sub

Public Property Get SearchNoteType() As Integer
    SearchNoteType = m_SearchNoteType
End Property

Public Property Let SearchNoteType(ByVal vNewValue As Integer)
    m_SearchNoteType = vNewValue

    If m_SearchNoteType = 8 Then
        Me.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì«     ÕÊÌ·«  „«·Ì…  "
        Me.XPLbl(3).Caption = "—ﬁ„ «·⁄„·Ì…"

        With Me.FG
            .TextMatrix(0, .ColIndex("NoteSerial")) = "—ﬁ„ «·⁄„·Ì…"
            .TextMatrix(0, .ColIndex("NoteDate")) = " «—ÌŒ «·⁄„·Ì…"
            .TextMatrix(0, .ColIndex("NoteValue")) = "ﬁÌ„… «·⁄„·Ì…"
        End With

    ElseIf m_SearchNoteType = 11 Or m_SearchNoteType = 12 Then
        Me.Caption = "«·»ÕÀ ⁄‰ «·“Ì«œ…  Ê«·‰ﬁ’ „‰ «·Œ“‰…"
        Me.XPLbl(3).Caption = "—ﬁ„ «·⁄„·Ì…"
        Me.CboType.Visible = True
        XPLbl(4).Visible = True

        With Me.FG
            .TextMatrix(0, .ColIndex("NoteSerial")) = "—ﬁ„ «·⁄„·Ì…"
            .TextMatrix(0, .ColIndex("NoteDate")) = " «—ÌŒ «·⁄„·Ì…"
            .TextMatrix(0, .ColIndex("NoteValue")) = "ﬁÌ„… «·⁄„·Ì…"
        End With

    End If

End Property

