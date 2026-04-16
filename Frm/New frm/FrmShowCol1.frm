VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowCol1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÙåÇÑ ÇáÃÚãÏÉ"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   Icon            =   "FrmShowCol1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   3465
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
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2715
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3315
      _cx             =   5847
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   18
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmShowCol1.frx":038A
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
      Default         =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ãæÇÝÞ"
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
      ButtonImage     =   "FrmShowCol1.frx":043A
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
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   609
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
      ButtonImage     =   "FrmShowCol1.frx":07D4
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÇáÃÚãÏÉ"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "FrmShowCol1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
 
Function ChangeLang()
 
    lbl(0).Caption = "Column Header"
    Me.Caption = "Show & Hide Columns"
 
    XPBtnOK.Caption = "&Ok"
    XPBtnCancel.Caption = "&Cancel"
 
End Function

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With FrmEmpSalary5.Grid

        Select Case Row

            Case 0

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Emp_Code")) = False
                Else
                    .ColHidden(.ColIndex("Emp_Code")) = True
                End If

            Case 1

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Emp_Name")) = False
                Else
                    .ColHidden(.ColIndex("Emp_Name")) = True
                End If

            Case 2

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Mokafea")) = False
                Else
                    .ColHidden(.ColIndex("Mokafea")) = True
                End If

            Case 3

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("TotalDiscount")) = False
                Else
                    .ColHidden(.ColIndex("TotalDiscount")) = True
                End If

            Case 4

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("TotalAdvance")) = False
                Else
                    .ColHidden(.ColIndex("TotalAdvance")) = False
                End If

            Case 5

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("total1")) = False
                Else
                    .ColHidden(.ColIndex("total1")) = True
                End If

            Case 6

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("total2")) = False
                Else
                    .ColHidden(.ColIndex("total2")) = True
                End If

            Case 7

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("EmpTotalNet")) = False
                Else
                    .ColHidden(.ColIndex("EmpTotalNet")) = True
                End If
            
            Case 8

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("sgn")) = False
                Else
                    .ColHidden(.ColIndex("sgn")) = True
                End If
            
        End Select

    End With

    With FrmEmpSalary5.Grid1

        Select Case Row

            Case 0

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Emp_Code")) = False
                Else
                    .ColHidden(.ColIndex("Emp_Code")) = True
                End If

            Case 1

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Emp_Name")) = False
                Else
                    .ColHidden(.ColIndex("Emp_Name")) = True
                End If

            Case 2

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("Mokafea")) = False
                Else
                    .ColHidden(.ColIndex("Mokafea")) = True
                End If

            Case 3

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("TotalDiscount")) = False
                Else
                    .ColHidden(.ColIndex("TotalDiscount")) = True
                End If

            Case 4

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("TotalAdvance")) = False
                Else
                    .ColHidden(.ColIndex("TotalAdvance")) = True
                End If

            Case 5

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("total1")) = False
                Else
                    .ColHidden(.ColIndex("total1")) = True
                End If

            Case 6

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("total2")) = False
                Else
                    .ColHidden(.ColIndex("total2")) = True
                End If

            Case 7

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("EmpTotalNet")) = False
                Else
                    .ColHidden(.ColIndex("EmpTotalNet")) = True
                End If
            
            Case 8

                If (FG.TextMatrix(Row, FG.ColIndex("show"))) = True Then
                    .ColHidden(.ColIndex("sgn")) = False
                Else
                    .ColHidden(.ColIndex("sgn")) = True
                End If
            
        End Select

    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
'On Error GoTo ErrTrap

'Dim i As Integer

Function fillgrigwithdata()
    Dim str(9) As String

    If SystemOptions.UserInterface = ArabicInterface Then
        str(1) = "ßæÏ ÇáãæÙÝ"
        str(2) = "ÇÓã ÇáãæÙÝ"
        str(3) = "ãßÇÝÃÊ  "
        str(4) = "ÌÒÇÁÇÊ"
        str(5) = "ÓáÝ"
        str(6) = "ÇÌãÇáí ÇáÇÖÇÝÇÊ  "
        str(7) = "ÇÌãÇáí ÇáÎÕæãÇÊ  "
        str(8) = "ÇáÕÇÝí  "
        str(9) = "ÇáÊæÞíÚ  "

    Else
        str(1) = "Emp Code "
        str(2) = " Emp Name"
        str(3) = "Bonus  "
        str(4) = "Punsh"
        str(5) = "Advance"
        str(6) = "Total Add "
        str(7) = "Total Dis "
        str(8) = "Net  "
        str(9) = "Sgn"

    End If

    Dim ColumnName1 As String
    Dim ColumnName2 As String
    Dim i As Integer

    With Me.FG
        .Rows = 9
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            '    .Rows = Rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To 9
                ColumnName1 = "S" & i
                ColumnName2 = "S" & i & "n"
             
                .TextMatrix(i - 1, .ColIndex("show")) = IIf(IsNull(rs.Fields(ColumnName1).value), "", rs.Fields(ColumnName1).value)
    
                .TextMatrix(i - 1, .ColIndex("name")) = str(i)
     
            Next i

            '           .TextMatrix(1, .ColIndex("show")) = IIf(IsNull(rs.Fields("s2").value), _
            '            "", rs.Fields("s2").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '                                   .TextMatrix(1, .ColIndex("name")) = IIf(IsNull(rs.Fields("s2n").value), _
            '            "", rs.Fields("s2n").value)
            '
            '            Else
            '             .TextMatrix(1, .ColIndex("name")) = IIf(IsNull(rs.Fields("s2ne").value), _
            '            "", rs.Fields("s2ne").value)
            '            End If
            
            '            .TextMatrix(2, .ColIndex("show")) = IIf(IsNull(rs.Fields("s3").value), _
            '            "", rs.Fields("s3").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(2, .ColIndex("name")) = IIf(IsNull(rs.Fields("s3n").value), _
            '            "", rs.Fields("s3n").value)
            '                   Else
            '             .TextMatrix(2, .ColIndex("name")) = IIf(IsNull(rs.Fields("s3ne").value), _
            '            "", rs.Fields("s3ne").value)
            '            End If
            
            '                .TextMatrix(3, .ColIndex("show")) = IIf(IsNull(rs.Fields("s4").value), _
            '            "", rs.Fields("s4").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(3, .ColIndex("name")) = IIf(IsNull(rs.Fields("s4n").value), _
            '            "", rs.Fields("s4n").value)
            '                   Else
            '             .TextMatrix(3, .ColIndex("name")) = IIf(IsNull(rs.Fields("s4ne").value), _
            '            "", rs.Fields("s4ne").value)
            '            End If
            
            '                        .TextMatrix(4, .ColIndex("show")) = IIf(IsNull(rs.Fields("s5").value), _
            '            "", rs.Fields("s5").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(4, .ColIndex("name")) = IIf(IsNull(rs.Fields("s5n").value), _
            '            "", rs.Fields("s5n").value)
            '                   Else
            '             .TextMatrix(4, .ColIndex("name")) = IIf(IsNull(rs.Fields("s5ne").value), _
            '            "", rs.Fields("s5ne").value)
            '            End If
            
            '          .TextMatrix(5, .ColIndex("show")) = IIf(IsNull(rs.Fields("s6").value), _
            '            "", rs.Fields("s6").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(5, .ColIndex("name")) = IIf(IsNull(rs.Fields("s6n").value), _
            '            "", rs.Fields("s6n").value)
            '                   Else
            '             .TextMatrix(5, .ColIndex("name")) = IIf(IsNull(rs.Fields("s6ne").value), _
            '            "", rs.Fields("s6ne").value)
            '            End If
            '            .TextMatrix(6, .ColIndex("show")) = IIf(IsNull(rs.Fields("s7").value), _
                         "", rs.Fields("s7").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(6, .ColIndex("name")) = IIf(IsNull(rs.Fields("s7n").value), _
            '            "", rs.Fields("s7n").value)
            '                   Else
            '             .TextMatrix(6, .ColIndex("name")) = IIf(IsNull(rs.Fields("s7ne").value), _
            '            "", rs.Fields("s7ne").value)
            '            End If
            '
            '         .TextMatrix(7, .ColIndex("show")) = IIf(IsNull(rs.Fields("s8").value), _
            '            "", rs.Fields("s8").value)
            
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(7, .ColIndex("name")) = IIf(IsNull(rs.Fields("s8n").value), _
            '            "", rs.Fields("s8n").value)
            '                   Else
            '             .TextMatrix(7, .ColIndex("name")) = IIf(IsNull(rs.Fields("s8ne").value), _
            '            "", rs.Fields("s8ne").value)
            '            End If
            '            .TextMatrix(8, .ColIndex("show")) = IIf(IsNull(rs.Fields("s9").value), _
            '            "", rs.Fields("s9").value)
            
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(8, .ColIndex("name")) = IIf(IsNull(rs.Fields("s9n").value), _
            '            "", rs.Fields("s9n").value)
            '                  Else
            '             .TextMatrix(8, .ColIndex("name")) = IIf(IsNull(rs.Fields("s9ne").value), _
            '            "", rs.Fields("s9ne").value)
            '            End If
            '           .TextMatrix(9, .ColIndex("show")) = IIf(IsNull(rs.Fields("s10").value), _
                        "", rs.Fields("s10").value)
            
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(9, .ColIndex("name")) = IIf(IsNull(rs.Fields("s10n").value), _
            '            "", rs.Fields("s10n").value)
            '                   Else
            '             .TextMatrix(9, .ColIndex("name")) = IIf(IsNull(rs.Fields("s10ne").value), _
            '            "", rs.Fields("s10ne").value)
            '            End If
            
            '           .TextMatrix(10, .ColIndex("show")) = IIf(IsNull(rs.Fields("s11").value), _
            '            "", rs.Fields("s11").value)
            '              If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(10, .ColIndex("name")) = IIf(IsNull(rs.Fields("s11n").value), _
            '            "", rs.Fields("s11n").value)
            ''                   Else
            '            .TextMatrix(10, .ColIndex("name")) = IIf(IsNull(rs.Fields("s11ne").value), _
            '           "", rs.Fields("s11ne").value)
            '           End If
            
            '           .TextMatrix(11, .ColIndex("show")) = IIf(IsNull(rs.Fields("s12").value), _
            '           "", rs.Fields("s12").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '           .TextMatrix(11, .ColIndex("name")) = IIf(IsNull(rs.Fields("s12n").value), _
            '           "", rs.Fields("s12n").value)
            '                  Else
            '            .TextMatrix(11, .ColIndex("name")) = IIf(IsNull(rs.Fields("s12ne").value), _
            '           "", rs.Fields("s12ne").value)
            '           End If
            '        .TextMatrix(12, .ColIndex("show")) = IIf(IsNull(rs.Fields("s13").value), _
            '           "", rs.Fields("s13").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '           .TextMatrix(12, .ColIndex("name")) = IIf(IsNull(rs.Fields("s13n").value), _
            '           "", rs.Fields("s13n").value)
            '                  Else
            '            .TextMatrix(12, .ColIndex("name")) = IIf(IsNull(rs.Fields("s13ne").value), _
            '           "", rs.Fields("s13ne").value)
            '           End If
            '          .TextMatrix(13, .ColIndex("show")) = IIf(IsNull(rs.Fields("s14").value), _
            '           "", rs.Fields("s14").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(13, .ColIndex("name")) = IIf(IsNull(rs.Fields("s14n").value), _
            '           "", rs.Fields("s14n").value)
            '                   Else
            '            .TextMatrix(13, .ColIndex("name")) = IIf(IsNull(rs.Fields("s14ne").value), _
            '           "", rs.Fields("s14ne").value)
            '           End If
            '            .TextMatrix(14, .ColIndex("show")) = IIf(IsNull(rs.Fields("s15").value), _
            '           "", rs.Fields("s15").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '           .TextMatrix(14, .ColIndex("name")) = IIf(IsNull(rs.Fields("s15n").value), _
            '           "", rs.Fields("s15n").value)
            '                  Else
            '            .TextMatrix(14, .ColIndex("name")) = IIf(IsNull(rs.Fields("s15ne").value), _
            '           "", rs.Fields("s15ne").value)
            '           End If
            '           .TextMatrix(15, .ColIndex("show")) = IIf(IsNull(rs.Fields("s16").value), _
            '           "", rs.Fields("s16").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '           .TextMatrix(15, .ColIndex("name")) = IIf(IsNull(rs.Fields("s16n").value), _
            '           "", rs.Fields("s16n").value)
            '                  Else
            '            .TextMatrix(15, .ColIndex("name")) = IIf(IsNull(rs.Fields("s16ne").value), _
            '           "", rs.Fields("s16ne").value)
            '           End If
            '           .TextMatrix(16, .ColIndex("show")) = IIf(IsNull(rs.Fields("s17").value), _
            '           "", rs.Fields("s17").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '           .TextMatrix(16, .ColIndex("name")) = IIf(IsNull(rs.Fields("s17n").value), _
            '           "", rs.Fields("s17n").value)
            '                  Else
            '            .TextMatrix(16, .ColIndex("name")) = IIf(IsNull(rs.Fields("s17ne").value), _
            '           "", rs.Fields("s17ne").value)
            '           End If
            '           .TextMatrix(17, .ColIndex("show")) = IIf(IsNull(rs.Fields("s18").value), _
            '           "", rs.Fields("s18").value)
            '             If SystemOptions.UserInterface = ArabicInterface Then
            '            .TextMatrix(17, .ColIndex("name")) = IIf(IsNull(rs.Fields("s18n").value), _
            '            "", rs.Fields("s18n").value)
            '                  Else
            '            .TextMatrix(17, .ColIndex("name")) = IIf(IsNull(rs.Fields("s18ne").value), _
            '           "", rs.Fields("s18ne").value)
            '           End If
            '
            '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
            '        Rs.MoveNext
            '    Next
            '   Rs.Close
        End If

        ' .RowHeight(-1) = 300
    End With

ErrTrap:
End Function

Private Sub Form_Load()

    Dim BGround As New ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    rs.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    fillgrigwithdata
    
    On Error GoTo ErrTrap
    Set FG.WallPaper = BGround.Picture
    CenterForm Me

    FormPostion Me, GetPostion
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnCancel_Click()
    'FrmMainPriceList.GetMeSetting
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    'On Error GoTo ErrTrap
    Dim i As Integer

    With FG

        For i = 0 To 8

            If .Cell(flexcpChecked, i, .ColIndex("show")) = flexChecked Then
                rs.Fields("s" & i + 1).value = True
            Else
                rs.Fields("s" & i + 1).value = False
            End If

        Next i
  
    End With

    rs.update
    'Rs.Close
    'Employee_salary_col
 
    Unload Me
    Exit Sub
ErrTrap:
End Sub
 
