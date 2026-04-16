VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchChangedCommponent 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·„›—œ«  Ê«·„ €Ì—« "
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmSearchChangedCommponent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   4350
      Width           =   4155
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3960
      Width           =   3195
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»"
      Height          =   645
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3240
      Width           =   10035
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ·"
         Height          =   195
         Index           =   3
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”«⁄« "
         Height          =   195
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ì«„"
         Height          =   195
         Index           =   1
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„…"
         Height          =   195
         Index           =   0
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2640
         TabIndex        =   31
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107085827
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107085827
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï  «—ÌŒ"
         Height          =   195
         Index           =   9
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   7
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2700
      Width           =   2115
      Begin VB.TextBox txtvalue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„Â"
         Height          =   195
         Index           =   4
         Left            =   1140
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»"
      Height          =   645
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5190
      Width           =   10035
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   1485
      End
      Begin VB.ComboBox CboYear 
         Height          =   315
         Left            =   3120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtname 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   3675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   195
         Index           =   8
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰Â"
         Height          =   195
         Index           =   0
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   5235
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   10035
      _cx             =   17701
      _cy             =   4842
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchChangedCommponent.frx":038A
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   5910
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
      Left            =   810
      TabIndex        =   7
      Top             =   5910
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
      Left            =   30
      TabIndex        =   8
      Top             =   5910
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
   Begin MSDataListLib.DataCombo DCComponent 
      Height          =   315
      Left            =   4800
      TabIndex        =   33
      Top             =   3960
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ« "
      Height          =   225
      Index           =   13
      Left            =   9000
      TabIndex        =   38
      Top             =   4470
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ã“¡ „‰ «”„ «·„›—œ"
      Height          =   195
      Index           =   12
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„›—œ"
      Height          =   195
      Index           =   11
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3960
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2940
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearchChangedCommponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim componentUnit As Integer

Private Sub CboYear_Change()
'  CboYear.text = year(XPDtbTrans.value)
'    CmbMonth.ListIndex = Month(XPDtbTrans.value) - 1
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    
 GetData
           
        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub Fg_Click()
FrmChangedComponentData.retrive1 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
'FrmChangedComponentData.Retrive1 val(Me.Fg.TextMatrix(Me.Fg.Row, Fg.ColIndex("id")))

End Sub


Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
YearMonth
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
     'Dcombos.GetClientName Me.DCEmp_Name
    '  If SystemOptions.UserInterface = EnglishInterface Then
     
    '  Me.DcbOrderStatus.AddItem "New"
    '    Me.DcbOrderStatus.AddItem "Accept Customer"
    '    Me.DcbOrderStatus.AddItem "Final Maintenance"

     '        Else
  Dim My_SQL As String
 If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " select id,name from mofrad where ViewComp=1 and  FixedOrChanged=1"
    Else
        My_SQL = " select id,namee from mofrad where ViewComp=1 and  FixedOrChanged=1"
    End If

    fill_combo DCComponent, My_SQL
  
' DcbOrderStatus.AddItem "ÃœÌœ"
'DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
'DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


'    End If
    Set DCboSearch = New clsDCboSearch
   ' Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
  SetDtpickerDate Me.DtpDateTo
  Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
Me.CboYear.ListIndex = -1
Me.CmbMonth.ListIndex = -1

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Private Sub DCComponent_Change()
'Dim Equation As Double
'    componentUnit = GetMofradUnit(val(Me.DCComponent.BoundText))
'
'    Opt(componentUnit).value = True
'    ChangeGridView componentUnit
'
    'LBLWhereSTR.Caption = GetComponentIncalculations(Val(Me.DCComponent.BoundText))
'    LBLWhereSTR.Caption = GetSpecificComponentIncalculations(val(Me.DCComponent.BoundText), Equation)
'LBLhOURrATE.Caption = Equation
    'v
End Sub
Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
   
    
StrSQL = " SELECT    dbo.TblChangedComponentRegister.RecordDate,dbo.TblChangedComponentRegister.Remarks , dbo.TblChangedComponentRegister.[year], dbo.TblChangedComponentRegister.[month], dbo.mofrad.id, "
     StrSQL = StrSQL & "                 dbo.mofrad.name, dbo.mofrad.nameE, dbo.mofrad.Absence, dbo.mofrad.Late, dbo.mofrad.Punch, dbo.mofrad.Discount, dbo.mofrad.OverTime,"
    StrSQL = StrSQL & "                  dbo.mofrad.AddOrDiscount, dbo.mofrad.Unit, dbo.mofrad.FixedOrChanged, dbo.mofrad.Account_Code, dbo.mofrad.ViewComp, dbo.mofrad.ZmamAccount,"
    StrSQL = StrSQL & "                  dbo.mofrad.Aloc1, dbo.mofrad.Aloc2, dbo.TblChangedComponentRegister.All_Or_SelectedEmployee, dbo.TblChangedComponentRegister.Actualyear,"
   StrSQL = StrSQL & "                   dbo.TblChangedComponentRegister.Actualmonth, dbo.TblChangedComponentRegister.BranchId, dbo.TblChangedComponentRegister.ChangedComponentid,"
   StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode,"
  StrSQL = StrSQL & "                     dbo.TblChangedComponentRegisterDetails.[value],"
  StrSQL = StrSQL & "                    dbo.TblChangedComponentRegisterDetails.GeneralChangedComponentid, dbo.TblChangedComponentRegisterDetails.HourRate,"
  StrSQL = StrSQL & "                    dbo.TblChangedComponentRegisterDetails.NoOfHour, dbo.TblChangedComponentRegisterDetails.NoOfMinutes, dbo.TblChangedComponentRegisterDetails.NoofDays,"
 StrSQL = StrSQL & "                     dbo.TblChangedComponentRegisterDetails.salary"
StrSQL = StrSQL & " FROM         dbo.mofrad INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblChangedComponentRegister ON dbo.mofrad.id = dbo.TblChangedComponentRegister.ComponentID INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblChangedComponentRegisterDetails ON"
 StrSQL = StrSQL & "                     dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID"
 '  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TBLSalesRepData2.BranchId = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData2.EmpID"
  
'StrSQL = "SELECT * FROM TblUnites "
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblChangedComponentRegister.ChangedComponentid >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.ChangedComponentid >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
   

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegister.ChangedComponentid <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.ChangedComponentid <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    '///////////////////
     If TxtValue.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegisterDetails.[value] =" & Me.TxtValue.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegisterDetails.[value] =" & Me.TxtValue.Text & ""
        End If
    End If
'////////////////////////
 If Me.txtname.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.txtname.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.txtname.Text & "%'"
        End If
    End If

If Me.TxtRemarks.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegister.Remarks like '%" & Me.TxtRemarks.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.Remarks like '%" & Me.TxtRemarks.Text & "%'"
        End If
    End If

       If CboYear.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegister.[year] =" & CboYear.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.[year] =" & CboYear.ListIndex & ""
        End If
    End If
      If DCComponent.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.mofrad.id =" & DCComponent.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.mofrad.id =" & DCComponent.BoundText & ""
        End If
    End If
    
        If CmbMonth.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegister.[month] =" & Me.CmbMonth.ListIndex & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.[month] =" & Me.CmbMonth.ListIndex & ""
        End If
    End If
 If Me.Opt(0).value = True Then
  If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.mofrad.Unit =0 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.mofrad.Unit =0 "
        End If
End If
 If Me.Opt(1).value = True Then
  If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.mofrad.Unit =1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.mofrad.Unit =1 "
        End If
End If
 If Me.Opt(2).value = True Then
  If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.mofrad.Unit =2 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.mofrad.Unit =2 "
        End If
End If
 If Me.Opt(3).value = True Then
  If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.mofrad.Unit >0 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.mofrad.Unit >0 "
        End If
End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblChangedComponentRegister.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblChangedComponentRegister.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblChangedComponentRegister.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblChangedComponentRegister.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
   End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblChangedComponentRegister.ChangedComponentid "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
               ' If rs("OrderStatus").value <> 2 Then
               ' Exit Sub
               ' End If
              
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ChangedComponentid").value), "", rs("ChangedComponentid").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                   .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Me.CboYear.ListIndex = IIf(IsNull(rs("year").value), -1, rs("year").value)
                .TextMatrix(i, .ColIndex("years")) = Me.CboYear.Text
                Me.CmbMonth.ListIndex = IIf(IsNull(rs("month").value), -1, rs("month").value)
                .TextMatrix(i, .ColIndex("months")) = Me.CmbMonth.Text
                .TextMatrix(i, .ColIndex("amount")) = IIf(IsNull(rs("value").value), "", rs("value").value)
                .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
               ' = IIf(IsNull(rs("Unit").value), "", rs("Unit").value) mofrdname
               .TextMatrix(i, .ColIndex("mofrdname")) = IIf(IsNull(rs("name").value), "", rs("name").value)
              If val(rs("Unit").value) = 0 Then
                .TextMatrix(i, .ColIndex("type")) = "ﬁÌ„Â"
                End If
                If val(rs("Unit").value) = 1 Then
                .TextMatrix(i, .ColIndex("type")) = "«Ì«„"
                End If
                 If val(rs("Unit").value) = 2 Then
                .TextMatrix(i, .ColIndex("type")) = "”«⁄« "
                End If
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    lbl(11).Caption = "Name Commponent"
    lbl(12).Caption = "Name Commponen"
  Me.Caption = "Search ChangedCommponent"
Me.Opt(0).RightToLeft = False
Me.Opt(0).Caption = "Value"
Me.Opt(1).RightToLeft = False
Me.Opt(1).Caption = "Daies"
Me.Opt(2).RightToLeft = False
Me.Opt(2).Caption = "Hours"
Me.Opt(0).RightToLeft = False
Me.Opt(3).Caption = "All"
'Me.LblClientName.Caption = "ClientName"
Me.Fra(0).Caption = "By"
Me.Fra(1).Caption = "By"
lbl(4).Caption = "Value"
lbl(3).Caption = "Emp Name"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(7).Caption = "From"
lbl(9).Caption = "To"
lbl(0).Caption = "Year"
lbl(8).Caption = "Month"
lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date "
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Process"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date Process"
        .TextMatrix(0, .ColIndex("code")) = "Code"
         .TextMatrix(0, .ColIndex("ClientName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("years")) = "Year"
       .TextMatrix(0, .ColIndex("months")) = "Month"
       .TextMatrix(0, .ColIndex("amount")) = "Value"
       .TextMatrix(0, .ColIndex("type")) = "Type"
        .TextMatrix(0, .ColIndex("mofrdname")) = "Name Commponen"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

