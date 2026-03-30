VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDiscountsSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ «·«‘⁄«—« "
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15285
   Icon            =   "FrmDiscountsSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TotalValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3240
      Width           =   1125
   End
   Begin VB.ComboBox CboDiscountType 
      Height          =   315
      Left            =   9780
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2490
      Width           =   2175
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2880
      Width           =   1125
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13140
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2490
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   223346689
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   345
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   223346689
         CurrentDate     =   38784
      End
      Begin VB.Label lBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   615
         Width           =   345
      End
      Begin VB.Label lBL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   2145
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   0
         Width           =   465
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   0
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   555
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1710
      TabIndex        =   11
      Top             =   4080
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
      Left            =   885
      TabIndex        =   12
      Top             =   4080
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
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   15135
      _cx             =   26696
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmDiscountsSearch.frx":038A
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
      Left            =   9840
      TabIndex        =   15
      Top             =   3750
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomers 
      Height          =   315
      Left            =   9840
      TabIndex        =   25
      Top             =   4170
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„Â »⁄œ ﬁ. „"
      Height          =   435
      Index           =   6
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   2
      Left            =   14160
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„Â ﬁ»· ﬁ.  „"
      Height          =   435
      Index           =   1
      Left            =   13980
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2910
      Width           =   1155
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ:"
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   315
      Index           =   3
      Left            =   14370
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2490
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·Œ’„"
      Height          =   315
      Index           =   4
      Left            =   12030
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
      Height          =   405
      Index           =   5
      Left            =   14160
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4110
      Width           =   1005
   End
End
Attribute VB_Name = "FrmDiscountsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_SearchType As Integer
Dim cSearchDcbo(4) As clsDCboSearch

Private Sub Cmd_Click(index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    On Error GoTo ErrTrap

    Select Case index

        Case 0
            Set rs = New ADODB.Recordset
            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
            XPLbl(0).Caption = "‰ ÌÃ… «·»ÕÀ :" & rs.RecordCount

            If rs.RecordCount < 1 Then
            
                fg.Clear flexClearScrollable, flexClearEverything
                fg.rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            Else

                With Me.fg
                    .Clear flexClearScrollable, flexClearEverything
                    .rows = .FixedRows
                    .rows = .FixedRows + rs.RecordCount
                    rs.MoveFirst

                    For i = .FixedRows To rs.RecordCount
                        .TextMatrix(i, .ColIndex("Serial")) = i
                        .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                        .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
.TextMatrix(i, .ColIndex("TotalValue")) = IIf(IsNull(rs("TotalValue").value), 0, rs("TotalValue").value)
.TextMatrix(i, .ColIndex("VATYou")) = IIf(IsNull(rs("VATYou").value), 0, rs("VATYou").value)
.TextMatrix(i, .ColIndex("VAT")) = IIf(IsNull(rs("VAT").value), 0, rs("VAT").value)


                        If Not IsNull(rs("NoteDate").value) Then
                            .TextMatrix(i, .ColIndex("NoteDate")) = Format(rs("NoteDate").value, "yyyy/M/d")
                        End If


 
 
 

       

                        If rs("NoteType").value = 9 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "Œ’„ „”„ÊÕ »Â"
                        ElseIf rs("NoteType").value = 10 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "Œ’„ „ﬂ ”»"
                              ElseIf rs("NoteType").value = 8034 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "œÌÊ‰ „⁄œÊ„…  "
                            
                              ElseIf rs("NoteType").value = 9082 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "«‘⁄«— „œÌ‰  "
                            
                                ElseIf rs("NoteType").value = 9083 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "«‘⁄«— œ«∆‰  "
                                 ElseIf rs("NoteType").value = 9089 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "ﬁÌ„… „÷«›…-«÷«›…  "
                            
                                      ElseIf rs("NoteType").value = 9090 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = "ﬁÌ„… „÷«›…-Œ’„  "
                            
                                      ElseIf rs("NoteType").value = 9099 Then
                            .TextMatrix(i, .ColIndex("PaymentType")) = " ’›Ì… «„·«ﬂ "
                        End If

                        .TextMatrix(i, .ColIndex("CustName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                        .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                        .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                    
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If
        
        Case 1
            clear_all Me
            fg.Clear flexClearScrollable, flexClearEverything
            fg.rows = 2

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub fg_Click()

    With Me.fg

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
    Dim GrdBack As New ClsBackGroundPic
    Set Dcombos = New ClsDataCombos
    CenterForm Me

    FormPostion Me, GetPostion
    Dcombos.GetUsers Me.DcboUsers
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboUsers
    Set cSearchDcbo(1) = New clsDCboSearch
   ' Set cSearchDcbo(1).Client = Me.DcboCustomers

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Me.CboDiscountType
        .Clear
        .AddItem "Œ’„ „”„ÊÕ »Â"
        .AddItem "Œ’„ „ﬂ ”»"
        .AddItem "  œÌÊ‰ „⁄œÊ„…"
        .AddItem "«‘⁄«— „œÌ‰"
        .AddItem "«‘⁄«— œ«∆‰"
        .AddItem " ﬁÌ„… „÷«›…-«÷«›…"
        .AddItem " ﬁÌ„… „÷«›…-Œ’„"
        .AddItem "«·ﬂ·"
    End With


    
    With Me.fg
        Set .WallPaper = GrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue
End Property

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim StrWhere As String

 '   StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.Remark, Notes.NoteHijriDate, Notes.CashingType," & "TblCustemers.CusID, TblCustemers.CusName, TblUsers.UserID, TblUsers.UserName"
 '   StrSQL = StrSQL + "  FROM TblUsers INNER JOIN (TblCustemers INNER JOIN Notes ON " & "TblCustemers.CusID = Notes.CusID) ON TblUsers.UserID = Notes.UserID "
StrSQL = "SELECT  dbo.Notes.TotalValue,dbo.Notes.VATYou,dbo.Notes.VAT,  dbo.Notes.NoteSerial1,   dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.Notes.Remark, "
 StrSQL = StrSQL + "                       dbo.Notes.NoteHijriDate , dbo.Notes.CashingType, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + "  FROM         dbo.TblCustemers RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                       dbo.TblUsers RIGHT OUTER JOIN"
 StrSQL = StrSQL + "                       dbo.Notes ON dbo.TblUsers.UserID = dbo.Notes.UserID ON dbo.TblCustemers.CusID = dbo.Notes.CusID "
                      
    If Me.CboDiscountType.ListIndex = -1 Or Me.CboDiscountType.ListIndex = 7 Then
        StrWhere = " Where (Notes.NoteType=9 OR Notes.NoteType=10  or Notes.NoteType=8034 or Notes.NoteType=9082 or Notes.NoteType=9083 or Notes.NoteType=9089 or Notes.NoteType=9090 or Notes.NoteType=9099)"
    ElseIf Me.CboDiscountType.ListIndex = 0 Then
        StrWhere = " Where (Notes.NoteType=9)"
    ElseIf Me.CboDiscountType.ListIndex = 1 Then
        StrWhere = " Where (Notes.NoteType=10)"
         ElseIf Me.CboDiscountType.ListIndex = 2 Then
        StrWhere = " Where (Notes.NoteType=8034)"
          ElseIf Me.CboDiscountType.ListIndex = 3 Then
       StrWhere = " Where (Notes.NoteType=9082)"
             
         ElseIf Me.CboDiscountType.ListIndex = 4 Then
       StrWhere = " Where (Notes.NoteType=9083)"
             
             
         ElseIf Me.CboDiscountType.ListIndex = 5 Then
       StrWhere = " Where (Notes.NoteType=9089)"
             
             
         ElseIf Me.CboDiscountType.ListIndex = 6 Then
       StrWhere = " Where (Notes.NoteType=9090)"
         ElseIf Me.CboDiscountType.ListIndex = 7 Then
       StrWhere = " Where (Notes.NoteType=9099)"
             
             
    End If

    If Trim(Me.TxtSerial.text) <> "" Then
        StrWhere = StrWhere + " AND NoteSerial1 =" & val(Me.TxtSerial.text) & ""
    End If

    If val(Me.TxtValue.text) > 0 Then
        If Me.opt(1).value = True Then
            StrWhere = StrWhere + " AND Notes.Note_Value =" & val(Me.TxtValue.text) & ""
        ElseIf Me.opt(0).value = True Then
            StrWhere = StrWhere + " AND Notes.Note_Value >" & val(Me.TxtValue.text) & ""
        Else
            StrWhere = StrWhere + " AND Notes.Note_Value <" & val(Me.TxtValue.text) & ""
        End If
    End If


    If val(Me.TotalValue.text) > 0 Then
        If Me.opt(1).value = True Then
            StrWhere = StrWhere + " AND Notes.TotalValue =" & val(Me.TotalValue.text) & ""
        ElseIf Me.opt(0).value = True Then
            StrWhere = StrWhere + " AND Notes.TotalValue >" & val(Me.TotalValue.text) & ""
        Else
            StrWhere = StrWhere + " AND Notes.TotalValue <" & val(Me.TotalValue.text) & ""
        End If
    End If
    
    
    If Me.DcboUsers.BoundText <> "" Then
        StrWhere = StrWhere + " AND Notes.UserID=" & Me.DcboUsers.BoundText & ""
    End If

    If Me.DcboCustomers.text <> "" Then
        StrWhere = StrWhere + " AND Notes.CusID=" & Me.DcboCustomers.BoundText & ""
    End If

    If Not IsNull(Me.DTPFrom.value) Then
        StrWhere = StrWhere + " AND  Notes.NoteDate >='" & SQLDate(Me.DTPFrom.value) & "'"
    End If
 
    If Not IsNull(Me.DTPTo.value) Then
        StrWhere = StrWhere + " AND  Notes.NoteDate <='" & SQLDate(Me.DTPTo.value) & "'"
    End If

    StrSQL = StrSQL + StrWhere + " Order By Notes.NoteID"
    Build_Sql = StrSQL
End Function

