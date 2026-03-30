VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSerialList 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÞÇÆãÉ ÇáÜ Serial Numbers ÇáãæÌæÏÉ"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "FrmSerialList.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÃÎÑ ÓÚÑ ÈíÚ"
      Height          =   1305
      Index           =   0
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3090
      Width           =   1275
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "0.00"
         Height          =   465
         Left            =   150
         TabIndex        =   24
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H000000FF&
      Caption         =   "ÇÓÚÇÑ ÇáÈíÚ ááÕäÝ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4470
      Width           =   4995
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÓÚÑ ÇáÏíáÑ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   510
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   210
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÓÚÑ ÇáÚãíá"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÓÚÑ ÇáãÓÊåáß"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   210
         Width           =   1155
      End
   End
   Begin ImpulseButton.ISButton CmdAdd 
      Height          =   285
      Left            =   390
      TabIndex        =   7
      Top             =   930
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   503
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSerialList.frx":038A
      DrawFocusRectangle=   0   'False
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   900
      Width           =   3225
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2715
      Left            =   30
      TabIndex        =   0
      Top             =   1710
      Width           =   3705
      _cx             =   6535
      _cy             =   4789
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSerialList.frx":0724
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
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5130
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÅáÛÇÁ"
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
      ButtonImage     =   "FrmSerialList.frx":07BC
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
   Begin ImpulseButton.ISButton cmdok 
      Height          =   375
      Left            =   870
      TabIndex        =   9
      Top             =   5130
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÍÝÙ"
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
      ButtonImage     =   "FrmSerialList.frx":0B56
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   345
      Index           =   0
      Left            =   3810
      TabIndex        =   13
      Top             =   1770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÊÍÏíÏ Çáßá"
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
      ButtonImage     =   "FrmSerialList.frx":0EF0
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   12582912
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   345
      Index           =   1
      Left            =   3810
      TabIndex        =   14
      Top             =   2160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ãÓÍ ÇáÊÍÏíÏ"
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
      ButtonImage     =   "FrmSerialList.frx":128A
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   12582912
   End
   Begin VB.Label LblItemID 
      Alignment       =   1  'Right Justify
      Height          =   225
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   60
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4830
      Picture         =   "FrmSerialList.frx":1624
      Top             =   1350
      Width           =   240
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇÏÎá ÓíÑíÇá ÇáÞØÚÉ ÇáãÈÇÚÉ æíãßäß Çä ÊÓÊÎÏã ÌåÇÒÇáÈÇÑßæÏ æ íãßäß Çä ÊÖÚ ÚáÇãÉ ÕÍ ÇãÇã ÇáÓíÑíÇá ÇáãÑÇÏ"
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
      Height          =   405
      Index           =   7
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1260
      Width           =   4755
   End
   Begin VB.Label LblStoreName 
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
      Height          =   255
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   4065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇÓã ÇáãÎÒä:"
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
      Height          =   255
      Index           =   6
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   5
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5340
      Width           =   555
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÅÌãÇáì ÇáßãíÉ ÇáãÎÊÇÑÉ:"
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
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5340
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5085
      X2              =   0
      Y1              =   5010
      Y2              =   5025
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   3
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5070
      Width           =   555
   End
   Begin VB.Label LblItemName 
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
      Height          =   225
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   330
      Width           =   4065
   End
   Begin VB.Label LblItemCode 
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
      Height          =   225
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇÓã ÇáÕäÝ:"
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
      Index           =   2
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   330
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ßæÏ ÇáÕäÝ:"
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
      Index           =   1
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÅÌãÇáì ÇáßãíÉ ÇáãæÌæÏÉ:"
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
      Index           =   0
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5070
      Width           =   1725
   End
End
Attribute VB_Name = "FrmSerialList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FgCol As Long

Public FgRow As Long

Public m_TextBox As TextBox

Public xGrid As ClsGrid

Private m_RetrunType As Integer

Private m_RetrunGridOnOneline As Boolean

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

            'Select All
            With Me.Fg
                .Cell(flexcpChecked, .FixedRows, .ColIndex("Selected"), .Rows - 1, .ColIndex("Selected")) = flexChecked
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Selected"), .Rows - 1, .ColIndex("Serial")) = vbGreen
                Me.lbl(5).Caption = GetSelectedCount
            End With

        Case 1

            'UnSelect All
            With Me.Fg
                .Cell(flexcpChecked, .FixedRows, .ColIndex("Selected"), .Rows - 1, .ColIndex("Selected")) = flexUnchecked
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Selected"), .Rows - 1, .ColIndex("Serial")) = 0
                Me.lbl(5).Caption = 0
            End With

    End Select

End Sub

Private Sub cmdAdd_Click()
    AddSerial
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim StrTemp As String
    Dim i As Integer
    Dim IntCheckCount As Integer
    Dim SngPrice As Single
    Dim IntCountAdd As Long

    IntCheckCount = GetFgCheckCount(Me.Fg, Fg.ColIndex("Selected"))

    If IntCheckCount = 0 Then
        Msg = "íÌÈ ÊÍÏíÏ ÓíÑíÇá æÇÍÏ Úáì ÇáÃÞá ..!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                StrTemp = StrTemp & .TextMatrix(i, .ColIndex("Serial")) & ";"
            End If

        Next i

    End With

    If Len(StrTemp) > 0 Then
        StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
    End If

    If Me.RetrunType = 0 Then
        If Me.RetrunGridOnOneline = True Then
            Me.xGrid.Grid.Cell(flexcpData, FgRow, FgCol) = StrTemp

            If IntCheckCount = 1 Then
                xGrid.Grid.TextMatrix(FgRow, FgCol) = StrTemp
            Else
                xGrid.Grid.TextMatrix(FgRow, FgCol) = "...."
            End If

            xGrid.Grid.TextMatrix(FgRow, m_Fg.ColIndex("Count")) = IntCheckCount
        Else

            '---------------------------------------------------------------------
            With Me.Fg
                IntCountAdd = Me.FgRow

                For i = .FixedRows To .Rows - 1

                    If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                        StrTemp = .TextMatrix(i, .ColIndex("Serial")) & ""
                        SngPrice = val(.TextMatrix(i, .ColIndex("Price")))
                        xGrid.AddItemInGrid IntCountAdd, val(Me.lblitemid.Caption), 1, SngPrice, , StrTemp
                        IntCountAdd = IntCountAdd + 1
                    End If

                Next i
        
            End With

        End If

    ElseIf Me.RetrunType = 1 Then
        Me.m_TextBox.text = StrTemp
    End If

    Unload Me
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With Me.Fg

        Select Case .ColKey(Col)

            Case "Selected"

                If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                    .Cell(flexcpBackColor, Row, .Col, Row, .ColIndex("Serial")) = vbGreen
                Else
                    .Cell(flexcpBackColor, Row, .Col, Row, .ColIndex("Serial")) = 0
                End If

        End Select

        Me.lbl(5).Caption = GetSelectedCount
    End With

End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    With Me.Fg

        Select Case .ColKey(Col)

            Case "Serial"
                Cancel = True

            Case Else
        
        End Select

    End With

End Sub

Private Sub Fg_DblClick()

    With Fg

        If .Col = -1 Then Exit Sub
        If .Row = -1 Then Exit Sub
        If .Col = .ColIndex("Serial") Then
            .Cell(flexcpChecked, .Row, .ColIndex("Selected")) = flexChecked
            FG_AfterEdit .Row, .ColIndex("Selected")
            CmdOk_Click
        End If

    End With

End Sub

Private Sub Fg_SelChange()
    Dim SngValue As Single
    Dim IntCheckState As VSFlex8UCtl.CellCheckedSettings

    If Fg.Col = Fg.ColIndex("Price") And Fg.ColSel = Fg.ColIndex("Price") Then
        If Trim(Fg.TextMatrix(Fg.Row, Fg.Col)) <> "" Then
            SngValue = val(Fg.TextMatrix(Fg.Row, Fg.Col))

            If SngValue > 0 Then
                Fg.Cell(flexcpText, Fg.Row, Fg.Col, Fg.RowSel, Fg.ColSel) = SngValue
            End If
        End If

    ElseIf Fg.Col = Fg.ColIndex("Selected") And Fg.ColSel = Fg.ColIndex("Selected") Then
        IntCheckState = Fg.Cell(flexcpChecked, Fg.Row, Fg.Col)
        Fg.Cell(flexcpChecked, Fg.Row, Fg.Col, Fg.RowSel, Fg.ColSel) = IntCheckState
    End If

End Sub

Private Sub Form_Load()
    Dim GrdBack As New ClsBackGroundPic

    CenterForm Me

    FormPostion Me, GetPostion
    Set Me.Fg.WallPaper = GrdBack.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then
        AddSerial
    End If

End Sub

Private Function GetSelectedCount() As Integer
    Dim IntCount As Integer
    Dim i As Integer

    With Me.Fg

        For i = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("Selected")) = flexChecked Then
                IntCount = IntCount + 1
            End If

        Next i

    End With

    GetSelectedCount = IntCount
End Function

Private Sub AddSerial()
    Dim StrSerial As String
    Dim LngFoundRow As Integer
    Dim Msg As String

    StrSerial = Trim(Me.TxtSerial.text)

    If StrSerial = "" Then Exit Sub

    With Me.Fg
        LngFoundRow = .FindRow(StrSerial, .FixedRows, .ColIndex("Serial"), False, True)

        If LngFoundRow <> -1 Then
            .Cell(flexcpChecked, LngFoundRow, .ColIndex("Selected")) = flexChecked
            FG_AfterEdit LngFoundRow, .ColIndex("Selected")
            Me.TxtSerial.text = ""
            Me.TxtSerial.SetFocus
        Else
            Msg = "åÐÇ ÇáÓíÑíÇá ÛíÑ ãæÌæÏ..."
            Msg = Msg & Chr(13) & "ÈÑÌÇÁ ÇáÊÇßÏ ãä ÇáÑÞã ÇáãÏÎá"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If

    End With

End Sub

Public Sub GetData(LngItemID As Long, _
                   LngStoreID As Long)

    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "select * From QryGardComplete"
        StrSQL = StrSQL + " where ItemID=" & LngItemID & ""
        StrSQL = StrSQL + " and StoreID =" & LngStoreID & ""
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT * From dbo.QryGardComplete(0)"
        StrSQL = StrSQL + " where ItemID=" & LngItemID & ""
        StrSQL = StrSQL + " and StoreID =" & LngStoreID & ""
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'Set Rs = GetItemQuantityStock(LngItemID, LngStoreID)
    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + rs.RecordCount

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Me.lblitemid.Caption = rs("ItemID").value
            Me.LblItemCode.Caption = rs("ItemCode").value
            Me.LblItemName.Caption = rs("ItemName").value
            Me.LblStoreName.Caption = rs("StoreName").value
    
            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("S")) = i
                .TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            
                rs.MoveNext
            Next i
        
        End If

        If .Rows > 1 Then
            Me.lbl(3).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("Serial"), .Rows - 1, .ColIndex("Serial"))
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtSerial_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
    'm_RetrunType= 0  Retrun into Grid
    'm_RetrunType= 1  Retrun into TextBOX
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    'm_RetrunType= 0  Retrun into Grid
    'm_RetrunType= 1  Retrun into TextBOX
End Property

Public Property Get RetrunGridOnOneline() As Boolean
    RetrunGridOnOneline = m_RetrunGridOnOneline
End Property

Public Property Let RetrunGridOnOneline(ByVal vNewValue As Boolean)
    m_RetrunGridOnOneline = vNewValue
End Property

Private Sub AfterEditSetting()
    Dim i As Long

    With Me.Fg

    End With

End Sub
