VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmGridColsShow 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌœ ŒÌ«—«  «·⁄—÷ ›Ï ÃœÊ· «·√’‰«›"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   ClipControls    =   0   'False
   Icon            =   "FrmGridColsShow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   5580
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
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Õ„Ì· «·ÕﬁÊ· «·«› —«÷ÌÂ"
      Height          =   375
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CheckBox Check17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ÕœÌœ / «·€«¡ «·ﬂ·"
      Height          =   195
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ÿ»Ìﬁ ⁄·Ï Ã„Ì⁄ «·‘«‘«  ›Ï «·»—‰«„Ã"
      Height          =   255
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5460
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   5145
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   5505
      _cx             =   9710
      _cy             =   9075
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
      BackColorFixed  =   -2147483633
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
      Rows            =   100
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmGridColsShow.frx":000C
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
      ExplorerBar     =   0
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
   Begin ImpulseButton.ISButton ISBXPBtnOK 
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      ButtonImage     =   "FrmGridColsShow.frx":00C9
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
   Begin ImpulseButton.ISButton ISBXPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   5880
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
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
      ButtonImage     =   "FrmGridColsShow.frx":0463
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5595
      X2              =   0
      Y1              =   5790
      Y2              =   5805
   End
End
Attribute VB_Name = "FrmGridColsShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FgGrid As VSFlex8UCtl.VSFlexGrid

Private Sub Check1_Click()
    Dim i As Integer


    If Check1.value = vbChecked Then

        With Me.Fg
 
            For i = 1 To .Rows - 1
        
     Debug.Print .TextMatrix(i, .ColIndex("ColKey"))
                If .TextMatrix(i, .ColIndex("ColKey")) = "Code" Or .TextMatrix(i, .ColIndex("ColKey")) = "Name" Or .TextMatrix(i, .ColIndex("ColKey")) = "Name" Or .TextMatrix(i, .ColIndex("ColKey")) = "Account_Name" _
                    Or .TextMatrix(i, .ColIndex("ColKey")) = "Account_Serial2" Or .TextMatrix(i, .ColIndex("ColKey")) = "UnitName" _
                Then
                           .TextMatrix(i, .ColIndex("ColShow")) = True
                
                End If
            Next i

        End With

    Else

        With Me.Fg

            For i = 1 To .Rows - 1
        
                Debug.Print .TextMatrix(i, .ColIndex("ColKey"))
            Next i

        End With

    End If


End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Fg
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ColShow")) = True
            Next i

        End With

    Else

        With Me.Fg

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ColShow")) = False
            Next i

        End With

    End If

 
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)

    With Me.Fg

        Select Case .ColKey(Col)

            Case "ColShow"
                Cancel = False

            Case Else
                Cancel = True
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    Set Me.Icon = Me.ISBXPBtnOK.ButtonImage
    CenterForm Me

    FormPostion Me, GetPostion

    With Me.Fg
        .Editable = flexEDKbdMouse
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Private Sub ChangeLang()
Check17.Caption = "Select All"
    Me.Caption = "Select  Column/s to Display"

    With Me.Fg

        .TextMatrix(0, .ColIndex("ColIndex")) = "I"
        .TextMatrix(0, .ColIndex("ColShow")) = "Show"
        .TextMatrix(0, .ColIndex("ColName")) = "Name"

    End With

    ISBXPBtnOK.Caption = "Ok"
    ISBXPBtnCancel.Caption = "Cancel"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub ISBXPBtnCancel_Click()

    Unload Me
End Sub

Private Sub ISBXPBtnOK_Click()
    Dim i As Integer
    Dim StrTemp As String
On Error Resume Next
    If Not FgGrid Is Nothing Then

        For i = 1 To Me.FgGrid.Cols - 1
            FgGrid.ColHidden(i) = True
        Next i
        
        For i = 1 To Me.Fg.Rows - 1
            FgGrid.ColHidden(FgGrid.ColIndex(Fg.TextMatrix(i, Fg.ColIndex("ColKey")))) = IIf(Fg.Cell(flexcpChecked, i, Fg.ColIndex("ColShow")) = flexChecked, False, True)
        Next i

    End If

    Unload Me
    'If Me.Chk.Value = vbChecked Then
    '    StrTemp = ""
    '    For I = 0 To Me.FgGrid.Cols - 1
    '        StrTemp = StrTemp + Me.FgGrid.ColKey(I) & "-" & Me.FgGrid.ColHidden(I) & ";"
    '    Next I
    '    StrTemp = Mid(StrTemp, 1, Len(StrTemp) - 1)
    '    '------------------------------------------
    '    SaveSetting SystemOptions.SysRegsAppPath, "GridOptions\ColsShow", _
    '    "FrmSaleBill", StrTemp
    '
    '    SaveSetting SystemOptions.SysRegsAppPath, "GridOptions\ColsShow", _
    '    "FrmBillBuy", StrTemp
    '
    '    SaveSetting SystemOptions.SysRegsAppPath, "GridOptions\ColsShow", _
    '    "FrmReturnpurchases", StrTemp
    '
    '    SaveSetting SystemOptions.SysRegsAppPath, "GridOptions\ColsShow", _
    '    "FrmReturnSalling", StrTemp
    '    '------------------------------------------
    'End If
End Sub
