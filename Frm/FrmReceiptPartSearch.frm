VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReceiptPartSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰  Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "FrmReceiptPartSearch.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   2130
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3480
      Width           =   1545
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   28
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
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   27
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
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   26
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄Ê«„· »ÕÀ ≈÷«›Ì…"
      ForeColor       =   &H00000080&
      Height          =   885
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5070
      Width           =   5595
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   1590
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì„ﬂ‰ﬂ «·»ÕÀ ⁄‰  Õ’Ì· «·√ﬁ”«ÿ «·Œ«’… »›« Ê—… »Ì⁄ «Ê ›« Ê—… ‘—«¡ „⁄Ì‰…"
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
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.ComboBox CboResType 
      Height          =   315
      Left            =   0
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2490
      Width           =   2205
   End
   Begin VB.TextBox TxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2940
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2820
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3450
      Width           =   2055
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   4210752
         CalendarTitleForeColor=   12648447
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   345
         Left            =   60
         TabIndex        =   8
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
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
         TabIndex        =   10
         Top             =   615
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
         TabIndex        =   9
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.ComboBox CboType 
      Height          =   315
      Left            =   2940
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2460
      Width           =   1635
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5715
      _cx             =   10081
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmReceiptPartSearch.frx":038A
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
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   2130
      TabIndex        =   4
      Top             =   3150
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   2100
      TabIndex        =   11
      Top             =   3870
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   2100
      TabIndex        =   14
      Top             =   4230
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   18
      Top             =   4620
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
      Left            =   825
      TabIndex        =   19
      Top             =   4620
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
      Left            =   60
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4620
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
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   4260
      TabIndex        =   21
      Top             =   4620
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ „ ﬁœ„..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmReceiptPartSearch.frx":0587
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmReceiptPartSearch.frx":0921
      ColorToggledHoverText=   192
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «· Õ’Ì·"
      Height          =   450
      Index           =   5
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2490
      Width           =   660
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   4
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4230
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   315
      Index           =   3
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3870
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· "
      Height          =   285
      Index           =   2
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3150
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   285
      Index           =   1
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2820
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   285
      Index           =   0
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2490
      Width           =   1005
   End
End
Attribute VB_Name = "FrmReceiptPartSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo(2) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim StrTemp As String

    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Set rs = New ADODB.Recordset
            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            Else

                With Me.FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = .FixedRows
                    .Rows = .FixedRows + rs.RecordCount
                    rs.MoveFirst
                
                    For i = .FixedRows To rs.RecordCount
                    
                        .TextMatrix(i, .ColIndex("Serial")) = i
                        .TextMatrix(i, .ColIndex("ReceiptID")) = IIf(IsNull(rs("ReceiptID").value), "", rs("ReceiptID").value)
                        .TextMatrix(i, .ColIndex("ReceiptID1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                     
                        If Not IsNull(rs("OperationType").value) Then
                            If rs("OperationType").value = 0 Then
                                .TextMatrix(i, .ColIndex("OperationType")) = " Õ’Ì· √ﬁ”«ÿ"
                            ElseIf rs("OperationType").value = 1 Then
                                .TextMatrix(i, .ColIndex("OperationType")) = "”œ«œ √ﬁ”«ÿ"
                            End If
                        End If

                        If Not IsNull(rs("ReceiptDate").value) Then
                            .TextMatrix(i, .ColIndex("ReceiptDate")) = DisplayDate(rs("ReceiptDate").value)
                        End If

                        If Not IsNull(rs("ReceiptType").value) Then
                            If rs("ReceiptType").value = 0 Then
                                .TextMatrix(i, .ColIndex("ReceiptType")) = " Õ’Ì· ⁄«œÏ"
                            ElseIf rs("ReceiptType").value = 1 Then
                                .TextMatrix(i, .ColIndex("ReceiptType")) = " Õ’Ì· »Œ’„"
                            ElseIf rs("ReceiptType").value = 2 Then
                                .TextMatrix(i, .ColIndex("ReceiptType")) = "œ›⁄… „‰ Õ”«» «·ﬁ”ÿ"
                            End If
                        End If

                        .TextMatrix(i, .ColIndex("PaymentMoney")) = IIf(IsNull(rs("PaymentMoney").value), "", rs("PaymentMoney").value)
                        .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                    
                        .TextMatrix(i, .ColIndex("Bankname")) = IIf(IsNull(rs("Bankname").value), "", IIf(IsNull(rs("Bankname").value), "", rs("Bankname").value))

                        If IsNull(rs("Bankname1").value) Then
                            .TextMatrix(i, .ColIndex("Bankname")) = ""
                        Else
                            .TextMatrix(i, .ColIndex("Bankname")) = rs("Bankname1").value
                        End If
                     
                        .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                        .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                        StrTemp = ""

                        If Not IsNull(rs("TransactionTypeName").value) Then
                            StrTemp = "„‰ Õ”«» ›« Ê—… " & rs("TransactionTypeName").value & " "
                        End If

                        If Not IsNull(rs("Transaction_Serial").value) Then
                            StrTemp = StrTemp & "—ﬁ„ " & rs("BillCode").value & " "
                        End If

                        .TextMatrix(i, .ColIndex("Notes")) = StrTemp
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub CmdShowMoreOptions_Click()

    If CmdShowMoreOptions.value = True Then
        Me.Fra(1).Visible = True
        Me.Height = Me.Fra(1).top + Fra(1).Height + 400
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Fra(1).top + 400
    
    End If

End Sub

Private Sub Fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ReceiptID"))) = 0 Then
            Exit Sub
        End If

        mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.Row, .ColIndex("ReceiptID")))
    End With

End Sub

Private Sub Form_Load()
    Dim cGrdBack As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos

    CenterForm Me

    FormPostion Me, GetPostion

    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DcboUsers
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboBox
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboUsers
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DBCboClientName

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Me.FG
        Set .WallPaper = cGrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.CboType
        .Clear
        .AddItem " Õ’Ì· √ﬁ”«ÿ"
        .AddItem "”œ«œ √ﬁ”«ÿ"
        .AddItem "«·ﬂ·"
    End With

    With CboResType
        .Clear
        .AddItem " Õ’Ì· ⁄«œÌ"
        .AddItem " Õ’Ì· »Œ’„"
        .AddItem "œ›⁄… „‰ Õ”«» «·ﬁ”ÿ"
        .AddItem "«·ﬂ·"
    End With

    With CboTrans
        .Clear
        .AddItem "›« Ê—… »Ì⁄"
        .AddItem "›« Ê—… „‘ —Ì« "
        .AddItem "«·ﬂ·"
    End With

    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Erase cSearchDcbo
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim StrWhere As String
    Dim IntTemp As Integer
    Dim BolBeginning As Boolean

    'StrSQL = "SELECT * From QryReceiptParts "

    StrSQL = "SELECT     dbo.ReceiptQest.ReceiptID, dbo.ReceiptQest.Cust_ID, dbo.ReceiptQest.ReceiptType, dbo.ReceiptQest.ReceiptDate, dbo.ReceiptQest.PartCount, dbo.ReceiptQest.Total, "
    StrSQL = StrSQL & "  dbo.ReceiptQest.PaymentMoney, dbo.ReceiptQest.User_ID, dbo.ReceiptQest.BoxID, dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName,"
    StrSQL = StrSQL & "   dbo.TblCustemers.CusName, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.TransactionTypes.TransactionTypeName,"
    StrSQL = StrSQL & "  dbo.ReceiptQest.OperationType, dbo.Transactions.Transaction_Type, dbo.ReceiptQest.NoteSerial, dbo.ReceiptQest.NoteSerial1, dbo.BanksData.BankName,"
    StrSQL = StrSQL & "  dbo.Transactions.NoteSerial1 AS BillCode, dbo.ReceiptQest.BankName AS Bankname1"
    StrSQL = StrSQL & "  FROM         dbo.TblBoxesData RIGHT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblUsers INNER JOIN"
    StrSQL = StrSQL & "  dbo.ReceiptQest INNER JOIN"
    StrSQL = StrSQL & "  dbo.Notes INNER JOIN"
    StrSQL = StrSQL & "   dbo.InstallMent INNER JOIN"
    StrSQL = StrSQL & "  dbo.InstallMentDetails INNER JOIN"
    StrSQL = StrSQL & "  dbo.InstallmentDet_Junc_Receipt ON dbo.InstallMentDetails.QestID = dbo.InstallmentDet_Junc_Receipt.QestID ON"
    StrSQL = StrSQL & "  dbo.InstallMent.PartID = dbo.InstallMentDetails.PartID ON dbo.Notes.NoteID = dbo.InstallMent.NoteID ON"
    StrSQL = StrSQL & "  dbo.ReceiptQest.ReceiptID = dbo.InstallmentDet_Junc_Receipt.ReceiptID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.ReceiptQest.Cust_ID = dbo.TblCustemers.CusID ON dbo.TblUsers.UserID = dbo.ReceiptQest.User_ID ON"
    StrSQL = StrSQL & "  dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type ON"
    StrSQL = StrSQL & "  dbo.TblBoxesData.BoxID = dbo.ReceiptQest.BoxID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.BanksData ON dbo.ReceiptQest.BankID = dbo.BanksData.BankID"
    BolBeginning = False

    If Me.CboType.ListIndex <> -1 And Me.CboType.ListIndex <> 2 Then
        If Me.CboType.ListIndex = 0 Then
            IntTemp = 0
        ElseIf Me.CboType.ListIndex = 1 Then
            IntTemp = 1
        End If

        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (OperationType=" & IntTemp & ")"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (OperationType=" & IntTemp & ")"
        End If
    End If

    If Me.CboResType.ListIndex <> -1 And Me.CboResType.ListIndex <> 3 Then
        If Me.CboResType.ListIndex = 0 Then
            IntTemp = 0
        ElseIf Me.CboResType.ListIndex = 1 Then
            IntTemp = 1
        ElseIf Me.CboResType.ListIndex = 2 Then
            IntTemp = 2
        End If

        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (ReceiptType=" & IntTemp & ")"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (ReceiptType=" & IntTemp & ")"
        End If
    End If

    If Trim(txtid.text) <> "" Then
        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (ReceiptQest.NoteSerial1='" & Trim(Me.txtid.text) & "')"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (ReceiptQest.NoteSerial1='" & Trim(Me.txtid.text) & "')"
        End If
    End If

    If Me.DBCboClientName.BoundText <> "" Then
        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (Cust_ID=" & Me.DBCboClientName.BoundText & ")"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (Cust_ID=" & Me.DBCboClientName.BoundText & ")"
        End If
    End If

    If Me.DcboBox.BoundText <> "" Then
        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (ReceiptQest.BoxID=" & Me.DcboBox.BoundText & ")"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (ReceiptQest.BoxID=" & Me.DcboBox.BoundText & ")"
        End If
    End If

    If Me.DcboUsers.BoundText <> "" Then
        If BolBeginning = False Then
            StrWhere = StrWhere & " Where (User_ID=" & Me.DcboUsers.BoundText & ")"
            BolBeginning = True
        Else
            StrWhere = StrWhere & " AND (User_ID=" & Me.DcboUsers.BoundText & ")"
        End If
    End If

    If val(Me.TxtValue.text) > 0 Then
        If BolBeginning = True Then
            If Me.Opt(1).value = True Then
                StrWhere = StrWhere + " AND (PaymentMoney =" & val(Me.TxtValue.text) & ")"
            ElseIf Me.Opt(0).value = True Then
                StrWhere = StrWhere + " AND (PaymentMoney >" & val(Me.TxtValue.text) & ")"
            Else
                StrWhere = StrWhere + " AND (PaymentMoney <" & val(Me.TxtValue.text) & ")"
            End If

        Else

            If Me.Opt(1).value = True Then
                StrWhere = StrWhere + " Where (PaymentMoney =" & val(Me.TxtValue.text) & ")"
            ElseIf Me.Opt(0).value = True Then
                StrWhere = StrWhere + " Where (PaymentMoney >" & val(Me.TxtValue.text) & ")"
            Else
                StrWhere = StrWhere + " Where (PaymentMoney <" & val(Me.TxtValue.text) & ")"
            End If

            BolBeginning = True
        End If
    End If

    'ReceiptDate
    If Not IsNull(Me.DTPFrom.value) Then
        If BolBeginning = True Then
            StrWhere = StrWhere & " AND (ReceiptDate >=" & SQLDate(Me.DTPFrom.value, True) & ")"
        Else
            StrWhere = StrWhere & " Where (ReceiptDate >=" & SQLDate(Me.DTPFrom.value, True) & ")"
            BolBeginning = True
        End If
    End If

    If Not IsNull(Me.DTPTo.value) Then
        If BolBeginning = True Then
            StrWhere = StrWhere & " AND (ReceiptDate <=" & SQLDate(Me.DTPTo.value, True) & ")"
        Else
            StrWhere = StrWhere & " Where (ReceiptDate <=" & SQLDate(Me.DTPTo.value, True) & ")"
            BolBeginning = True
        End If
    End If

    If CmdShowMoreOptions.value = True Then
        If Me.CboTrans.ListIndex <> -1 And Me.CboTrans.ListIndex <> 2 Then
            If Me.CboTrans.ListIndex = 0 Then
                IntTemp = 21
            ElseIf Me.CboTrans.ListIndex = 1 Then
                IntTemp = 22
            End If

            If BolBeginning = True Then
                StrWhere = StrWhere & " AND (Transactions.Transaction_Type =" & IntTemp & ")"
            Else
                StrWhere = StrWhere & " Where (Transactions.Transaction_Type=" & IntTemp & ")"
                BolBeginning = True
            End If

            If Trim(Me.TxtTransSerial.text) <> "" Then
                If BolBeginning = True Then
                    StrWhere = StrWhere & " AND (dbo.Transactions.NoteSerial1='" & Trim(Me.TxtTransSerial.text) & "')"
                Else
                    StrWhere = StrWhere & " Where (dbo.Transactions.NoteSerial1='" & Trim(Me.TxtTransSerial.text) & "')"
                    BolBeginning = True
                End If
            End If
        End If
    End If

    Build_Sql = StrSQL + StrWhere
End Function
