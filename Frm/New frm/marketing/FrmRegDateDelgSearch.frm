VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRegDateDelgSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·» „Ê«⁄Ìœ «·„‰«œÌ»"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14340
   Icon            =   "FrmRegDateDelgSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   14340
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”«⁄…"
      Height          =   675
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   5520
      Width           =   3975
      Begin MSComCtl2.DTPicker Timeto 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   93650946
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   315
         Left            =   2040
         TabIndex        =   42
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   93650946
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   14
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   15
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   675
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Width           =   4815
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2370
         TabIndex        =   22
         Top             =   270
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93650947
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   270
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93650947
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   9
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   7
         Left            =   1815
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   645
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2640
      Width           =   6435
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   5655
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1485
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   14235
      Begin VB.TextBox DcbJobID 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5400
         TabIndex        =   43
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox TxtMobi 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox TxtTel 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5400
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TxtAdmini 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   3135
      End
      Begin MSDataListLib.DataCombo DcbTypeVisit1 
         Height          =   315
         Left            =   10080
         TabIndex        =   19
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   10080
         TabIndex        =   20
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbJobID1 
         Height          =   315
         Left            =   4440
         TabIndex        =   28
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbTypeVisit2 
         Height          =   315
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCustomer 
         Height          =   315
         Left            =   10080
         TabIndex        =   44
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ŒÿÊ… «· «·ÌÂ"
         Height          =   195
         Index           =   13
         Left            =   3585
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«·"
         Height          =   285
         Index           =   19
         Left            =   3480
         TabIndex        =   35
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ì„Ì·"
         Height          =   285
         Index           =   20
         Left            =   3480
         TabIndex        =   34
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰"
         Height          =   285
         Index           =   18
         Left            =   8670
         TabIndex        =   31
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÊŸÌ›… «·„”ƒ·"
         Height          =   195
         Index           =   12
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”ƒ·"
         Height          =   195
         Index           =   11
         Left            =   9210
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·“Ì«—…"
         Height          =   195
         Index           =   4
         Left            =   13095
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   195
         Index           =   8
         Left            =   13530
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‰œÊ»"
         Height          =   195
         Index           =   0
         Left            =   13095
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1020
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14355
      _cx             =   25321
      _cy             =   4630
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmRegDateDelgSearch.frx":038A
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
      Left            =   2490
      TabIndex        =   1
      Top             =   4800
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
      Left            =   1650
      TabIndex        =   2
      Top             =   4800
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
      Left            =   870
      TabIndex        =   3
      Top             =   4800
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊŸ›"
      Height          =   195
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2940
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   -780
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2820
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   -660
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmRegDateDelgSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

 GetData
           
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.TimeFrom.value = ""
Me.Timeto.value = ""
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
     'On Error GoTo ErrTrap
  '  FrmModels.FindRec
   FrmRegDateDelgate.Retrive val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))
'ErrTrap:
                
            
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
'GetData
  Dcombos.GetTypeVisit Me.DcbTypeVisit1
    Dcombos.GetTypeVisit Me.DcbTypeVisit2
    'Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetDelegate Me.DcboEmpName
   Dcombos.GetFileCustomer Me.DcbCustomer
    'Dcombos.GetEmpJobsTypes Me.DcbJobID
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
   DtpDateFrom.value = ""
   Me.DtpDateTo.value = ""
Me.TimeFrom.value = ""
Me.Timeto.value = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = "SELECT DISTINCT "
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Name1,"
 StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4, TblEmployee_1.Nationality, TblEmployee_1.Emp_Namee1,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee3, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee4, dbo.TblRegDateDelgate.Id, TblEmployee_1.Emp_ID,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.BranchID, dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name AS Emp_NameD, TblEmployee_1.Emp_Name1 AS Emp_Name1D, TblEmployee_1.Emp_Name2 AS Emp_Name2D,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name3 AS Emp_Name3D, TblEmployee_1.Emp_Name4 AS Emp_Name4D, TblEmployee_1.Nationality AS NationalityD,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee AS Emp_NameeD, TblEmployee_1.Emp_Namee1 AS Emp_Namee1D, TblEmployee_1.Emp_Namee2 AS Emp_Namee2D,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee3 AS Emp_Namee3D, TblEmployee_1.Emp_Namee4 AS Emp_Namee4D, TblEmployee_1.Fullcode AS FullcodeD,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID, dbo.TblTypeVisit.name, dbo.TblTypeVisit.namee,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.VisitID2, TblTypeVisit_1.name AS name2, TblTypeVisit_1.namee AS namee2, dbo.TblRegDateDelgate.SpAsID,"
StrSQL = StrSQL & "                       dbo.TblSpeciaAsement.name AS nameSp, dbo.TblSpeciaAsement.namee AS nameeSp, dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.Remark2, dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.LongTime, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, dbo.TblRegDateDelgate.CustomerID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                       dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.FromTime1, dbo.TblRegTimeDelgate.name AS FromTime11, dbo.TblRegDateDelgate.FromTime2,"
StrSQL = StrSQL & "                       TblRegTimeDelgate_1.name AS FromTime22, dbo.TblRegDateDelgate.ToTime2 AS ToTime2, TblRegTimeDelgate_3.name AS ToTime22,"
 StrSQL = StrSQL & "                      dbo.TblRegDateDelgate.ToTime1, TblRegTimeDelgate_2.name AS ToTime11"
StrSQL = StrSQL & "  FROM         dbo.TblRegTimeDelgate TblRegTimeDelgate_1 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_3 RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegDateDelgate LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_2 ON dbo.TblRegDateDelgate.ToTime1 = TblRegTimeDelgate_2.Id ON"
  StrSQL = StrSQL & "                     TblRegTimeDelgate_3.Id = dbo.TblRegDateDelgate.ToTime2 ON TblRegTimeDelgate_1.Id = dbo.TblRegDateDelgate.FromTime2 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate ON dbo.TblRegDateDelgate.FromTime1 = dbo.TblRegTimeDelgate.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblSpeciaAsement ON dbo.TblRegDateDelgate.SpAsID = dbo.TblSpeciaAsement.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID2 = TblTypeVisit_1.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblTypeVisit ON dbo.TblRegDateDelgate.VisitID = dbo.TblTypeVisit.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID"
    BolBegine = False
    StrWhere = ""
Dim str, str1 As String
If Not IsNull(Me.TimeFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblRegDateDelgate.TimeFrom1 <=" & Format(Me.TimeFrom.value, "hh:mm ") & ""
        Else
            BolBegine = True
            StrWhere = " Where   dbo.TblRegDateDelgate.TimeFrom1 <=" & Format(Me.TimeFrom.value, "hh:mm") & ""
        End If
    End If
   If Me.TimeFrom.value <> "" Then
   'str = Me.TimeFrom.value
        If BolBegine = True Then
 
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.TimeFrom1 >=" & Format(Me.TimeFrom.value, "hh:mm AM/PM") & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.TimeFrom1 >=" & val(str) & ""
        End If
    End If
 
 
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.Id >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.Id >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
 

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.Id <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.Id<=" & val(Me.TxtIDTO.text) & ""
        End If
    End If


'////////////////////////
' If Me.txtname.text <> "" Then
'        If BolBegine = True Then
'            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.CustomerName like '%" & Me.txtname.text & "%'"
'        Else
'            BolBegine = True
'            StrWhere = " Where dbo.TblRegDateDelgate.CustomerName like '%" & Me.txtname.text & "%'"
'        End If
'    End If
    
   ''''
      If Me.TxtTel.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.Tel like '%" & Me.TxtTel.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.Tel like '%" & Me.TxtTel.text & "%'"
        End If
    End If
   ''
         If Me.TxtMobi.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.Mobile like '%" & Me.TxtMobi.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.Mobile like '%" & Me.TxtMobi.text & "%'"
        End If
    End If
    
           If Me.TxtEmail.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.Email like '%" & Me.TxtEmail.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.Email like '%" & Me.TxtEmail.text & "%'"
        End If
    End If
   '''
   If Me.TxtAdmini.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.PersonConc like '%" & Me.TxtAdmini.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.PersonConc like '%" & Me.TxtAdmini.text & "%'"
        End If
    End If
 If Me.DcboEmpName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblEmployee.Emp_ID =" & DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_ID=" & DcboEmpName.BoundText & ""
        End If
    End If
    ''
     If Me.DcbJobID.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.JobID like '%" & Me.DcbJobID.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.JobID like '%" & Me.DcbJobID.text & "%'"
        End If
    End If
    ''
    If Me.DcbTypeVisit2.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.VisitID2 =" & DcbTypeVisit2.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.VisitID2=" & DcbTypeVisit2.BoundText & ""
        End If
    End If
 If Me.DcbCustomer.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.CustomerID =" & DcbCustomer.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.CustomerID=" & DcbCustomer.BoundText & ""
        End If
    End If
 If Me.DcbTypeVisit1.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblRegDateDelgate.VisitID =" & DcbTypeVisit1.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblRegDateDelgate.VisitID=" & DcbTypeVisit1.BoundText & ""
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblRegDateDelgate.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND   dbo.TblRegDateDelgate.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where   dbo.TblRegDateDelgate.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblRegDateDelgate.id"
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
       MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
             
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
       If SystemOptions.UserInterface = EnglishInterface Then
           .TextMatrix(i, .ColIndex("delgate")) = IIf(IsNull(rs("Emp_NameeD").value), "", rs("Emp_NameeD").value)
            .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
           Else
          .TextMatrix(i, .ColIndex("delgate")) = IIf(IsNull(rs("Emp_NameD").value), "", rs("Emp_NameD").value)
            .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
           End If
          
          '  .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("admini")) = IIf(IsNull(rs("PersonConc").value), "", rs("PersonConc").value)
                  If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("recorddate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                     .TextMatrix(i, .ColIndex("job")) = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
            .TextMatrix(i, .ColIndex("tel")) = IIf(IsNull(rs("Tel").value), "", rs("Tel").value)
                .TextMatrix(i, .ColIndex("mobile")) = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
                .TextMatrix(i, .ColIndex("time1")) = val(IIf(IsNull(rs("FromTime11").value), "", rs("FromTime11").value))
                .TextMatrix(i, .ColIndex("timto")) = val(IIf(IsNull(rs("ToTime11").value), "", rs("ToTime11").value))
                     .TextMatrix(i, .ColIndex("email")) = IIf(IsNull(rs("Email").value), "", rs("Email").value)
            .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                .TextMatrix(i, .ColIndex("type2")) = IIf(IsNull(rs("name2").value), "", rs("name2").value)
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
  Me.Caption = "Search End_Service"
lbl(4).Caption = "Type"
'Me.LblClientName.Caption = "AsestName"
Frame1.Caption = "Reg Time"
lbl(14).Caption = "From"
lbl(15).Caption = "To"
lbl(0).Caption = "Delegate"
lbl(8).Caption = "Customer"
lbl(2).Caption = "Total"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbreg.Caption = "Reg Date"
lbl(9).Caption = "From"
lbl(7).Caption = "To"
lbl(11).Caption = "Admin"
lbl(12).Caption = "Job"
lbl(4).Caption = "Type Visit"
lbl(18).Caption = "Phone"
lbl(19).Caption = "Mobile"
lbl(20).Caption = "Email"
lbl(13).Caption = "Next Step"
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No"
        .TextMatrix(0, .ColIndex("recorddate")) = "Date"
         .TextMatrix(0, .ColIndex("time1")) = "From Time"
        .TextMatrix(0, .ColIndex("delgate")) = "Delgate"
      .TextMatrix(0, .ColIndex("admini")) = "Admin"
             .TextMatrix(0, .ColIndex("ClientName")) = "Customer"
        .TextMatrix(0, .ColIndex("job")) = "Job"
        .TextMatrix(0, .ColIndex("tel")) = "Phone"
         .TextMatrix(0, .ColIndex("mobile")) = "Mobile"
        .TextMatrix(0, .ColIndex("email")) = "Email"
      .TextMatrix(0, .ColIndex("type")) = "TypeVisit"
      .TextMatrix(0, .ColIndex("type2")) = "Next Step"
       .TextMatrix(0, .ColIndex("timto")) = "To Time"
    End With
  '
End Sub

Private Sub TimeFrom1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

