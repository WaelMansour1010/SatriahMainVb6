VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSup_ItemsPrice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘—«∆Õ √”⁄«— «·’‰›"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "FrmSup_ItemsPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox XPTxtPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   900
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox TxtCompareValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   555
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   1965
      Left            =   150
      TabIndex        =   1
      Top             =   1410
      Width           =   3525
      _cx             =   6218
      _cy             =   3466
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
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSup_ItemsPrice.frx":038A
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
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   3435
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmSup_ItemsPrice.frx":0415
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Height          =   405
      Left            =   120
      TabIndex        =   5
      Top             =   3435
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
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
      ButtonImage     =   "FrmSup_ItemsPrice.frx":07AF
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnRemove 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3450
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmSup_ItemsPrice.frx":0B49
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton XPBtnAdd 
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   3450
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmSup_ItemsPrice.frx":0EE3
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label XPLblPriceID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1170
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label XPLblItemName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   405
      Width           =   1935
   End
   Begin VB.Label XPLblSupName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‘—«∆Õ «·√”⁄«—"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   2580
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”⁄— «·‘—«¡"
      Height          =   255
      Index           =   4
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   255
      Index           =   6
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   405
      Width           =   885
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ê—œ"
      Height          =   255
      Index           =   7
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "FrmSup_ItemsPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim OldGrdValue As String

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Dim Msg As String
    On Error GoTo ErrTrap

    With FG

        If .TextMatrix(Row, Col) <> "" Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Or Len((.TextMatrix(Row, Col))) > 50 Then
                Msg = "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ…" & Chr(13)
                Msg = Msg + "√œŒ· ﬁÌ„ —ﬁ„Ì… „ÊÃ»… "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                .TextMatrix(Row, Col) = OldGrdValue
                Exit Sub
            ElseIf .TextMatrix(Row, Col) < 0 Then
                Msg = "√œŒ· ﬁÌ„ —ﬁ„Ì… „ÊÃ»… "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                .TextMatrix(Row, Col) = OldGrdValue
                Exit Sub
            End If
        End If

        If Col = .ColIndex("To") Then
            If val(.TextMatrix(Row, Col)) < val(.TextMatrix(Row, .ColIndex("Form"))) Then
                Msg = "·«»œ √‰ ÌﬂÊ‰ «·Õœ «·√ﬁ’ ··‘—ÌÕ…(≈·Ï)" & Chr(13)
                Msg = Msg + "√ﬂ»— „‰ «·Õœ «·√œ‰Ï ·Â« („‰) "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If

        If Col = .ColIndex("Form") Then
            If Row > 1 Then
                If val(.TextMatrix(Row, Col)) - val(.TextMatrix(Row - 1, .ColIndex("To"))) > 1 Or val(.TextMatrix(Row, Col)) - val(.TextMatrix(Row - 1, .ColIndex("To"))) < 1 Then
                    Msg = "ÌÃ» √‰  »œ√ «·‘—ÌÕ… «·ÃœÌœ… " & Chr(13)
                    Msg = Msg + "„‰ ÕÌÀ «‰ Â  «·‘—ÌÕ… «·”«»ﬁ… +1"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

        If Col = .ColIndex("Price") Then
            If Row >= 1 Then
                If val(.TextMatrix(Row, .ColIndex("Price"))) > val(XPTxtPrice.text) Then
                    Msg = "√”⁄«— «·‘—«∆Õ ·«»œ √‰  ﬂÊ‰ " & Chr(13)
                    Msg = Msg + "√ﬁ· „‰ ”⁄— «·»Ì⁄"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            If Row > 1 Then
                If val(.TextMatrix(Row, .ColIndex("Price"))) > val(.TextMatrix(Row - 1, .ColIndex("Price"))) Then
                    Msg = "ÌÃ» √‰ ÌﬂÊ‰ ”⁄— «·‘—ÌÕ… √ﬁ·" & Chr(13)
                    Msg = Msg + "„‰ ”⁄— «·‘—ÌÕ… «·”«»ﬁ…" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Form")) <> "" And FG.TextMatrix(FG.Rows - 1, FG.ColIndex("To")) <> "" And FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Price")) <> "" Then
            FG.Rows = FG.Rows + 1
            FG.TextMatrix(FG.Rows - 1, FG.ColIndex("NumIndex")) = FG.Rows - 1
            FG.Row = FG.Rows - 1
            FG.Col = FG.ColIndex("Form")
        
        Else
            FG.Row = FG.Rows - 1
            FG.Col = FG.ColIndex("Form")
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    StrSQL = "select * From JuncPrice where juncID=" & Me.XPLblPriceID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FG.TextMatrix(1, FG.ColIndex("NumIndex")) = "1"

    If Not (rs.EOF Or rs.BOF) Then
        Retrive
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyF2 Then
        XPBtnAdd_Click
    End If

    If KeyCode = vbKeyF3 Then
        XPBtnRemove_Click
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BackGround As ClsBackGroundPic
    Set BackGround = New ClsBackGroundPic
    FG.WallPaper = BackGround.MoneyWallpaper
    CenterForm Me

    FormPostion Me, GetPostion
    AddTip

    With FG
        .Cell(flexcpPicture, 0, .ColIndex("Form")) = mdifrmmain.ImgLstTree.ListImages("From").Picture
        .Cell(flexcpPicture, 0, .ColIndex("To")) = mdifrmmain.ImgLstTree.ListImages("To").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Price")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnAdd_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Form")) <> "" Then
        FG.Rows = FG.Rows + 1
        FG.TextMatrix(FG.Rows - 1, FG.ColIndex("NumIndex")) = FG.Rows - 1
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim RsItems As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RsPrice As ADODB.Recordset
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim ColNum As Integer
    Dim BeginTrans As Boolean

    Dim Msg As String

    If Not IsNumeric(XPTxtPrice.text) Then
        Msg = "”⁄— «·‘—«¡ «·–Ì √œŒ· Â €Ì— ’«·Õ" & Chr(13)
        Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    With FG

        For RowNum = 1 To .Rows - 1
            For ColNum = 1 To .Cols - 1

                If .TextMatrix(RowNum, ColNum) <> "" Then
                    If Not IsNumeric(.TextMatrix(RowNum, ColNum)) Or Len((.TextMatrix(RowNum, ColNum))) > 50 Then
                        Msg = "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ…" & Chr(13)
                        Msg = Msg + " √ﬂœ √‰ Ã„Ì⁄ «·ﬁÌ„ —ﬁ„Ì… "
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If

            Next ColNum

            If val(.TextMatrix(RowNum, .ColIndex("To"))) < val(.TextMatrix(RowNum, .ColIndex("Form"))) Then
                Msg = "·«»œ √‰ ÌﬂÊ‰ «·Õœ «·√ﬁ’ ··‘—ÌÕ…(≈·Ï)" & Chr(13)
                Msg = Msg + "√ﬂ»— „‰ «·Õœ «·√œ‰Ï ·Â« („‰) " & Chr(13)
                Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If RowNum > 1 Then
                If .TextMatrix(RowNum, .ColIndex("Form")) <> "" Then
                    If val(.TextMatrix(RowNum, .ColIndex("Form"))) - val(.TextMatrix(RowNum - 1, .ColIndex("To"))) > 1 Or val(.TextMatrix(RowNum, .ColIndex("Form"))) - val(.TextMatrix(RowNum - 1, .ColIndex("To"))) < 1 Then
                        Msg = "ÌÃ» √‰  »œ√ ﬂ· ‘—ÌÕ… " & Chr(13)
                        Msg = Msg + "„‰ ÕÌÀ «‰ Â  «·‘—ÌÕ… «·”«»ﬁ… +1" & Chr(13)
                        Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If

                If val(.TextMatrix(RowNum, .ColIndex("Price"))) > val(.TextMatrix(RowNum - 1, .ColIndex("Price"))) Then
                    Msg = "ÌÃ» √‰ ÌﬂÊ‰ ”⁄— «·‘—ÌÕ… √ﬁ·" & Chr(13)
                    Msg = Msg + "„‰ ”⁄— «·‘—ÌÕ… «·”«»ﬁ…" & Chr(13)
                    Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            If RowNum >= 1 Then
                If val(.TextMatrix(RowNum, .ColIndex("Price"))) > val(XPTxtPrice.text) Then
                    Msg = "√”⁄«— «·‘—«∆Õ ·«»œ √‰  ﬂÊ‰ " & Chr(13)
                    Msg = Msg + "√ﬁ· „‰ ”⁄— «·‘—«¡ «·«› —«÷Ì"
                    Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

        Next RowNum

    End With

    StrSQL = "delete * From JuncPrice where juncID=" & Me.XPLblPriceID
    Cn.Execute StrSQL
    StrSQL = "select * From CusJuncItem where ID=" & Me.XPLblPriceID
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Cn.BeginTrans
    BeginTrans = True
    RsItems("ItemPrice").value = Me.XPTxtPrice
    RsItems("LastUpdate").value = Date

    With FrmMainPriceList.FgMain
        .TextMatrix(.Row, .ColIndex("DefalutPrice")) = Me.XPTxtPrice

        If val(TxtCompareValue.text) <> val(XPTxtPrice.text) Then
            .TextMatrix(.Row, .ColIndex("LastUpdate")) = Format(Date, "yyyy/mm/dd")
        End If

    End With

    RsItems.update

    For RowNum = 1 To FG.Rows - 1

        With FG

            If .TextMatrix(RowNum, .ColIndex("Price")) <> "" Then
                rs.AddNew
                rs("PriceID").value = new_id("JuncPrice", "PriceID", "", True)
                rs("juncID").value = Me.XPLblPriceID
                rs("From").value = IIf(.TextMatrix(RowNum, .ColIndex("Form")) = "", "0", val(.TextMatrix(RowNum, .ColIndex("Form"))))
                rs("to").value = IIf(IsNull(.TextMatrix(RowNum, .ColIndex("To"))), "0", .TextMatrix(RowNum, .ColIndex("To")))
                rs("Price").value = IIf(IsNull(.TextMatrix(RowNum, .ColIndex("Price"))), "0", .TextMatrix(RowNum, .ColIndex("Price")))
                rs.update
            End If

        End With

    Next RowNum

    Cn.CommitTrans
    BeginTrans = False
    ' „ÌÌ“ «·√’‰«› «· Ì ·Â« ‘—«∆Õ √”⁄«—
    Set RsPrice = New ADODB.Recordset

    If Me.XPLblPriceID <> "" Then
        StrSQL = "select * From JuncPrice where juncID=" & Me.XPLblPriceID
        RsPrice.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsPrice.EOF Or RsPrice.BOF) Then

            With FrmMainPriceList.FgMain
                .Cell(flexcpPicture, .Row, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
            End With

        Else

            With FrmMainPriceList.FgMain
                .Cell(flexcpPicture, .Row, .ColIndex("DefalutPrice")) = ""
            End With

        End If

        RsPrice.Close
    End If

    Unload Me
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.Rows > 1 Then
        If FG.Rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
        Else

            If FG.Rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Rows - 1)
                End If
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    FG.Rows = rs.RecordCount + 1

    For RowNum = 1 To rs.RecordCount

        With FG
            .TextMatrix(RowNum, .ColIndex("NumIndex")) = RowNum
            .TextMatrix(RowNum, .ColIndex("Form")) = IIf(IsNull(rs("From").value), "", Trim(rs("From").value))
            .TextMatrix(RowNum, .ColIndex("To")) = IIf(IsNull(rs("To").value), "", Trim(rs("To").value))
            .TextMatrix(RowNum, .ColIndex("Price")) = IIf(IsNull(rs("Price").value), "", Trim(rs("Price").value))
        End With

        rs.MoveNext
    Next RowNum

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "‘—«∆Õ √”⁄«— „Ê—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnAdd, "≈÷«›… ‘—ÌÕ… ..." & Wrap & "·«÷«›… ‘—ÌÕ… √”⁄«— ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘—«∆Õ √”⁄«— „Ê—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnRemove, "Õ–› ‘—ÌÕ… ..." & Wrap & "·Õ–› ‘—ÌÕ… «·√”⁄«— «·√ŒÌ—…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘—«∆Õ √”⁄«— „Ê—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPTxtPrice, "”⁄— «·‘—«¡ «·«› —«÷Ì ··’‰›", True
    End With

    With TTP
        .Create Me.hWnd, "‘—«∆Õ √”⁄«— „Ê—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnOK, "„Ê«›ﬁ ..." & Wrap & "·Õ›Ÿ ‘—«∆Õ «·√”⁄«— «· Ì  „ ﬂ «» Â«" & Wrap & " √Ê «· ⁄œÌ·«  «· Ì  „  ⁄·ÌÂ«", True
    End With

    With TTP
        .Create Me.hWnd, "‘—«∆Õ √”⁄«— „Ê—œ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnCancel, "≈·€«¡  ..." & Wrap & "·≈·€«¡ «· ⁄œÌ·«  «· Ì  „  ⁄·Ï ‘—«∆Õ «·√”⁄«—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
End Sub

