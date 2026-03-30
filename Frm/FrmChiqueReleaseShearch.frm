VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChiqueReleaseShearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰  Õ’Ì· Ê”œ«œ «·‘Ìﬂ« "
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "FrmChiqueReleaseShearch.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   2070
      TabIndex        =   6
      Top             =   4050
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo2"
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox CboOperaType 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2460
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄Ê«„· »ÕÀ ≈÷«›Ì…"
      ForeColor       =   &H00000080&
      Height          =   1335
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4950
      Width           =   5775
      Begin VB.ComboBox CboCheckType 
         Height          =   315
         Left            =   2520
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2025
      End
      Begin MSDataListLib.DataCombo DcboBanks 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpNoteDate 
         Height          =   345
         Left            =   2940
         TabIndex        =   13
         Top             =   570
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DtpDueDate 
         Height          =   345
         Left            =   2940
         TabIndex        =   14
         Top             =   930
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·√” Õﬁ«ﬁ"
         Height          =   315
         Index           =   9
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· Õ—Ì—"
         Height          =   315
         Index           =   8
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   315
         Index           =   4
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   210
         Width           =   645
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·‘Ìﬂ"
         Height          =   315
         Index           =   0
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.TextBox TxtOperaID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3630
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2460
      Width           =   1395
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3630
      Width           =   1545
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
         TabIndex        =   24
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   555
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
         TabIndex        =   23
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   0
         Value           =   -1  'True
         Width           =   495
      End
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
         TabIndex        =   22
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   945
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2820
      Width           =   1935
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
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
         Top             =   570
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   615
         Width           =   345
      End
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2520
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2850
      Width           =   2505
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3630
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3630
      Width           =   1395
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _cx             =   10398
      _cy             =   4260
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmChiqueReleaseShearch.frx":038A
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
      Index           =   0
      Left            =   1770
      TabIndex        =   15
      Top             =   4500
      Width           =   825
      _ExtentX        =   1455
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
      TabIndex        =   16
      Top             =   4500
      Width           =   855
      _ExtentX        =   1508
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4500
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   30
      TabIndex        =   9
      Top             =   4050
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomers 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   3240
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   4470
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
      ButtonImage     =   "FrmChiqueReleaseShearch.frx":056C
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmChiqueReleaseShearch.frx":0906
      ColorToggledHoverText=   192
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «”„ «·Œ“‰…"
      Height          =   315
      Index           =   11
      Left            =   5070
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4050
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      Height          =   315
      Index           =   7
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   2460
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   315
      Index           =   6
      Left            =   5070
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2490
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ ⁄«„·"
      Height          =   315
      Index           =   5
      Left            =   5070
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·‘Ìﬂ"
      Height          =   315
      Index           =   3
      Left            =   5070
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   5070
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3660
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” Œœ„ «·„Õ——"
      Height          =   195
      Index           =   2
      Left            =   690
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3810
      Width           =   1305
   End
End
Attribute VB_Name = "FrmChiqueReleaseShearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim m_SearchType As Integer
Dim cSearchDcbo(4) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Set rs = New ADODB.Recordset
            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

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
                        .TextMatrix(i, .ColIndex("OperaID")) = IIf(IsNull(rs("OperaID").value), "", rs("OperaID").value)
                        .TextMatrix(i, .ColIndex("OperaDate")) = IIf(IsNull(rs("OperaDate").value), "", rs("OperaDate").value)
                        .TextMatrix(i, .ColIndex("OperaType")) = IIf(IsNull(rs("CheckTypeName").value), "", rs("CheckTypeName").value)
                        .TextMatrix(i, .ColIndex("ChqueNum")) = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
                        .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                        .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("NoteValue").value), "", rs("NoteValue").value)
                        .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                        .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
                        .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)
                        .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If

        Case 1
            clear_all Me
            SetClearDates
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
        Me.Height = Me.Fra(1).top + Me.Fra(1).Height + 600
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Cmd(0).top + Me.Cmd(0).Height + 600
    End If

    'Me.Height = Me.ScaleHeight + 500
End Sub

Private Sub Fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("OperaID"))) = 0 Then
            Exit Sub
        End If

        mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.Row, .ColIndex("OperaID")))
    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As New ClsBackGroundPic

    Set Dcombos = New ClsDataCombos
    CenterForm Me

    FormPostion Me, GetPostion
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers, False
    Dcombos.GetBanks Dcbobanks
    Dcombos.GetUsers Me.DcboUsers
    Dcombos.GetBoxes Me.DcboBox
    'Dcombos.GetUsers Me.DcboIussedUser

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboCustomers
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.Dcbobanks

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboUsers

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Me.CboCheckType
        .Clear
        .AddItem "‘Ìﬂ ··‘—ﬂ…"
        .AddItem "‘Ìﬂ ⁄·Ï «·‘—ﬂ…"
        .AddItem "‘Ìﬂ „— œ ⁄·Ï ⁄„Ì· «Ê „Ê—œ"
        .AddItem "‘Ìﬂ „— œ ⁄·Ï «·‘—ﬂ…"
        .AddItem "«·ﬂ·"
    End With

    With Me.CboOperaType
        .Clear
        .AddItem " Õ’Ì· ‘Ìﬂ ··‘—ﬂ…"
        .AddItem "”œ«œ ‘Ìﬂ ⁄·Ï «·‘—ﬂ…"
        .AddItem "«·ﬂ·"
    End With

    With Me.FG
        Set .WallPaper = GrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    SetClearDates

    CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub TxtOperaID_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtOperaID.text, 1)
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue

    With Me.FG

        If m_SearchType = 3 Then
            ' 3 «·»ÕÀ ⁄‰ «·„’—Ê›« 
            .ColHidden(.ColIndex("PaymentType")) = False
            .ColHidden(.ColIndex("CustName")) = True
            Me.Caption = "«·»ÕÀ ⁄‰ «·„’—Ê›« "
            Me.XPLbl(4).Visible = True

            Me.XPLbl(5).Visible = False
            Me.DcboCustomers.Visible = False
        ElseIf m_SearchType = 4 Then
            ' 4 «·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« 
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.Caption = "«·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« "
            Me.XPLbl(4).Visible = False

            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True

        ElseIf m_SearchType = 5 Then
            '5     «·»ÕÀ «·„œ›Ê⁄« 
            Me.Caption = "«·»ÕÀ ⁄‰ «·„œ›Ê⁄« "
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.XPLbl(4).Visible = False

            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True
        End If

    End With

End Property

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim StrWhere As String

    StrSQL = "Select * From QryCheckRelease "

    If Me.CboCheckType.ListIndex = -1 Or Me.CboCheckType.ListIndex = 2 Then
        StrSQL = StrSQL + " Where (NoteType = 2 Or NoteType = 13)"
    ElseIf Me.CboCheckType.ListIndex = 0 Then
        StrSQL = StrSQL + " Where (NoteType = 2)"
    ElseIf Me.CboCheckType.ListIndex = 1 Then
        StrSQL = StrSQL + " Where (NoteType = 13)"
    End If

    If Trim(Me.TxtOperaID.text) <> "" Then
        StrWhere = StrWhere + " AND OperaID=" & val(Me.TxtOperaID.text) & ""
    End If

    If Me.CboOperaType.ListIndex = -1 Or Me.CboOperaType.ListIndex = 2 Then
        StrSQL = StrSQL + " And (OperaType = 0 OR OperaType=1 OR OperaType=2 OR OperaType=3)"
    ElseIf Me.CboOperaType.ListIndex = 0 Then
        StrSQL = StrSQL + " And (OperaType = 0)"
    ElseIf Me.CboOperaType.ListIndex = 1 Then
        StrSQL = StrSQL + " And (OperaType = 1)"
    ElseIf Me.CboOperaType.ListIndex = 2 Then
        StrSQL = StrSQL + " And (OperaType = 2)"
    ElseIf Me.CboOperaType.ListIndex = 3 Then
        StrSQL = StrSQL + " And (OperaType = 3)"
    End If

    If val(Me.TxtValue.text) > 0 Then
        If Me.Opt(1).value = True Then
            StrWhere = StrWhere + " AND NoteValue =" & val(Me.TxtValue.text) & ""
        ElseIf Me.Opt(0).value = True Then
            StrWhere = StrWhere + " AND NoteValue >" & val(Me.TxtValue.text) & ""
        Else
            StrWhere = StrWhere + " AND NoteValue <" & val(Me.TxtValue.text) & ""
        End If
    End If

    If Trim(TxtSerial.text) <> "" Then
        StrWhere = StrWhere + " AND ChqueNum Like '%" & Trim(TxtSerial.text) & "%'"
    End If

    If Me.DcboUsers.BoundText <> "" Then
        StrWhere = StrWhere + " AND UserID=" & Me.DcboUsers.BoundText & ""
    End If

    If Me.DcboCustomers.BoundText <> "" Then
        StrWhere = StrWhere + " AND CusID=" & Me.DcboCustomers.BoundText & ""
    End If

    If Me.Dcbobanks.BoundText <> "" Then
        StrWhere = StrWhere + " AND  BankID=" & Me.Dcbobanks.BoundText & ""
    End If

    If Me.DcboBox.BoundText <> "" Then
        StrWhere = StrWhere + " AND  BoxID=" & Me.DcboBox.BoundText & ""
    End If

    If Not IsNull(Me.DTPFrom.value) Then
        StrWhere = StrWhere + " AND  OperaDate >=#" & SQLDate(Me.DTPFrom.value) & "#"
    End If

    If Not IsNull(Me.DTPTo.value) Then
        StrWhere = StrWhere + " AND  OperaDate <=#" & SQLDate(Me.DTPTo.value) & "#"
    End If

    If Not IsNull(Me.DtpNoteDate.value) Then
        StrWhere = StrWhere + " AND  NoteDate =#" & SQLDate(Me.DtpNoteDate.value) & "#"
    End If

    If Not IsNull(Me.DTPDueDate.value) Then
        StrWhere = StrWhere + " AND  DueDate=#" & SQLDate(Me.DTPDueDate.value) & "#"
    End If

    StrSQL = StrSQL + StrWhere + " Order By OperaID"

    Build_Sql = StrSQL
End Function

Private Sub SetClearDates()
    SetDtpickerDate Me.DTPDueDate
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DtpNoteDate
    SetDtpickerDate Me.DTPTo

End Sub
