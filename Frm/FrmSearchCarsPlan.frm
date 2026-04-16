VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchCarsPlan 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ Œÿ… ’Ì«‰… «·„—ﬂ»« "
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "FrmSearchCarsPlan.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   8820
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
   Begin VB.ComboBox CboType1 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      ItemData        =   "FrmSearchCarsPlan.frx":038A
      Left            =   0
      List            =   "FrmSearchCarsPlan.frx":038C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Õ—ﬂ…"
      Height          =   645
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   5235
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   2460
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4575
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·› —Â"
      Height          =   1215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3120
      Width           =   3255
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   118030339
         CurrentDate     =   41640
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   118030339
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï  «—ÌŒ"
         Height          =   195
         Index           =   2
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1575
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   5505
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   720
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.TextBox TXTCurrentKM 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo Dccar 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCMaintenanceTypes 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·’Ì«‰Â"
         Height          =   315
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ—«¡… «·Õ«·Ì… ··⁄œ«œ"
         Height          =   315
         Index           =   12
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„⁄œÂ/«·”Ì«—…"
         Height          =   285
         Index           =   3
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   -60
      TabIndex        =   0
      Top             =   270
      Width           =   8835
      _cx             =   15584
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchCarsPlan.frx":038E
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
      TabIndex        =   1
      Top             =   4440
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
      TabIndex        =   2
      Top             =   4440
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
      TabIndex        =   3
      Top             =   4440
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
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   2310
      Left            =   -30
      TabIndex        =   29
      Top             =   30
      Width           =   8790
      _cx             =   15505
      _cy             =   4075
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   42
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchCarsPlan.frx":0487
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
      ExplorerBar     =   3
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   2295
   End
End
Attribute VB_Name = "FrmSearchCarsPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Indx As Integer

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            If Indx = 3 Then
                GetData2
            Else
                GetData
            End If
           
        Case 1
            clear_all Me
            DtpDateFrom.value = ""
DtpDateTo.value = ""
'Me.DtpDateFrom.value = ""
'Me.DtpDateTo.value = ""
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
If Me.Indx = 1 Then
FrmCarExpireLicens.txtid.Text = val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
Else
FrmCarsPlan.Retrive val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
End If
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  Set Dcombos = New ClsDataCombos
    Dcombos.GetEquipments Me.Dccar
    Dcombos.GetCarsMaintenanceTypes Me.DCMaintenanceTypes

    
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
    GridInstallments.Visible = False
    If Indx = 3 Then
        Fg.Visible = False
        GridInstallments.Visible = True
        DCMaintenanceTypes.Visible = False
        lbl(0).Visible = False
        TXTCurrentKM.Visible = False
        lbl(12).Visible = False
    End If
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
DtpDateFrom.value = ""
DtpDateTo.value = ""

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

StrSQL = "SELECT     dbo.TblCarMaintenancePlan.Planid, dbo.TblCarMaintenancePlan.RecordDate, dbo.TblCarMaintenancePlan.CurrentKM, dbo.TblCarMaintenancePlan.CarId, "
StrSQL = StrSQL & "                      dbo.TblCarsData.BoardNO, dbo.TblCarMaintenancePlanDetails.MaintenanceID, dbo.MaintenanceTypes.NAME, dbo.MaintenanceTypes.km,"
StrSQL = StrSQL & "                       dbo.TblCarMaintenancePlan.PlanYear , dbo.MaintenanceTypes.NameE"
StrSQL = StrSQL & "  FROM         dbo.MaintenanceTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarMaintenancePlanDetails ON dbo.MaintenanceTypes.id = dbo.TblCarMaintenancePlanDetails.MaintenanceID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarMaintenancePlan ON dbo.TblCarMaintenancePlanDetails.Planid = dbo.TblCarMaintenancePlan.Planid LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarsData ON dbo.TblCarMaintenancePlan.CarId = dbo.TblCarsData.id"

    BolBegine = False
    StrWhere = ""

    '///////////////////
        If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCarMaintenancePlan.Planid >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.Planid >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.Planid <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.Planid <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    

    

    
          If (Me.Dccar.Text <> "") And (val(Dccar.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarsData.fixedAssetid=" & Me.Dccar.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarsData.fixedAssetid =" & Me.Dccar.BoundText & ""
        End If
    End If
 ''//
     If (Me.DCMaintenanceTypes.Text) <> "" And (val(DCMaintenanceTypes.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlanDetails.MaintenanceID =" & Me.DCMaintenanceTypes.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlanDetails.MaintenanceID =" & Me.DCMaintenanceTypes.BoundText & ""
        End If
    End If
 ''//
 
 
       If (Me.TXTCurrentKM.Text) <> "" And (val(TXTCurrentKM.Text) <> 0) Then
       If Me.Opt(2).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.CurrentKM < " & val(Me.TXTCurrentKM.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.CurrentKM < " & val(Me.TXTCurrentKM.Text) & ""
        End If
         ElseIf Me.Opt(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.CurrentKM =" & val(Me.TXTCurrentKM.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.CurrentKM =" & val(Me.TXTCurrentKM.Text) & ""
        End If
           ElseIf Me.Opt(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.CurrentKM > " & val(Me.TXTCurrentKM.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.CurrentKM > " & val(Me.TXTCurrentKM.Text) & ""
        End If
     
        End If
        
    End If
     If Not IsNull(Me.DtpDateFrom.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.TblCarMaintenancePlan.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
                   
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarMaintenancePlan.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.TblCarMaintenancePlan.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
                   
      End If


    '-----------------------------------
  StrWhere = StrWhere & "  GROUP BY dbo.TblCarMaintenancePlan.Planid, dbo.TblCarMaintenancePlan.RecordDate, dbo.TblCarMaintenancePlan.CurrentKM, dbo.TblCarMaintenancePlan.CarId,"
  StrWhere = StrWhere & "                    dbo.TblCarsData.BoardNO, dbo.TblCarMaintenancePlanDetails.MaintenanceID, dbo.MaintenanceTypes.NAME, dbo.MaintenanceTypes.km,"
  StrWhere = StrWhere & "                    dbo.TblCarMaintenancePlan.PlanYear , dbo.MaintenanceTypes.NameE "
                      

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblCarMaintenancePlan.Planid "
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
               
                .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("Planid").value), "", rs("Planid").value)
                .TextMatrix(i, .ColIndex("CurrentKM")) = IIf(IsNull(rs("CurrentKM").value), "", rs("CurrentKM").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("NAME").value), "", rs("NAME").value)
                    Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                    End If
         

                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub


Public Sub GetData2()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer



 StrSQL = " SELECT   tblTripTrans.Id as ID2,tblTripTrans.recordDate,tblTripTrans.recordDateH,tblTripTrans.Fromdate,tblTripTrans.todate,"
StrSQL = StrSQL & "                       notesallid , dbo.tblTripTrans2.notesallid, dbo.tblTripTrans2.ID, dbo.tblTripTrans2.TravID, dbo.tblTripTrans2.TripNo, dbo.tblTripTrans2.TripDate, dbo.tblTripTrans2.BranchID, "
StrSQL = StrSQL & "                      TblBoxesData.BoxName , Accounts.account_name,Branches2.branch_name,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.Price,dbo.tblTripTrans2.TotalValue,tblTripTrans2.RecNo,tblTripTrans2.Weight,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblTripTrans2.Typed, dbo.tblTripTrans2.[Value], dbo.tblTripTrans2.Remarks,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.NoteID, dbo.tblTripTrans2.QtyDownload, dbo.tblTripTrans2.QtyDischarge, dbo.tblTripTrans2.CardNO, dbo.tblTripTrans2.CardNO2,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarType1, dbo.tblTripTrans2.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.tblTripTrans2.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.tblTripTrans2.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.tblTripTrans2.TypeTrans, dbo.tblTripTrans2.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.tblTripTrans2.LeaderName,TblCustemers.CusName ,TblCustemers.CusID ,TblCarsData.BoardNO"
StrSQL = StrSQL & " FROM         dbo.tblTripTrans2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.tblTripTrans2.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.tblTripTrans2.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.tblTripTrans2.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.tblTripTrans2.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.tblTripTrans2.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.tblTripTrans2.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblTripTrans2.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblCustemers On TblCustemers.CusID = dbo.tblTripTrans2.CusID "
StrSQL = StrSQL & "                      LEFT OUTER JOIN tblTripTrans On tblTripTrans.Id = tblTripTrans2.TravID "
StrSQL = StrSQL & "                      LEFT OUTER JOIN ACCOUNTS On tblTripTrans.AccountPaym = ACCOUNTS.Account_Code "
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblBranchesData Branches2 On tblTripTrans.BranchId = Branches2.branch_id "
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblBoxesData  On tblTripTrans.BoxID = TblBoxesData.BoxID "





    BolBegine = False
    StrWhere = ""

    '///////////////////
        If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.tblTripTrans.id >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.tblTripTrans.id >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblTripTrans.id <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarMaintenancePlan.id <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
    

    

    
          If (Me.Dccar.Text <> "") And (val(Dccar.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCarsData.fixedAssetid=" & Me.Dccar.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCarsData.fixedAssetid =" & Me.Dccar.BoundText & ""
        End If
    End If
 ''//
  
 ''//
 
 
     If Not IsNull(Me.DtpDateFrom.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblTripTrans.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.tblTripTrans.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
                   
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
       If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.tblTripTrans.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
          StrWhere = StrWhere & " where dbo.tblTripTrans.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
                   
      End If


'    '-----------------------------------
'  StrWhere = StrWhere & "  GROUP BY dbo.tblTripTrans.id, dbo.tblTripTrans.RecordDate, dbo.TblCarMaintenancePlan.CarId,"
'  StrWhere = StrWhere & "                    dbo.TblCarsData.BoardNO, dbo.TblCarMaintenancePlanDetails.MaintenanceID, dbo.MaintenanceTypes.NAME, dbo.MaintenanceTypes.km,"
'  StrWhere = StrWhere & "                    dbo.TblCarMaintenancePlan.PlanYear , dbo.MaintenanceTypes.NameE "
'
   ' StrWhere = StrWhere & "   Where  (dbo.tblTripTrans2.TypeTrans is null or dbo.tblTripTrans2.TypeTrans=0) and "
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.tblTripTrans.id "
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

        With Me.GridInstallments
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
                If SystemOptions.UserInterface = ArabicInterface Then
                 .ColComboList(.ColIndex("Show")) = "⁄—÷"
                Else
                .ColComboList(.ColIndex("Show")) = "View"
                End If
                
                .TextMatrix(i, .ColIndex("Ser")) = i
                   .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
                   'RsDetails1("notesallid").value = val(.TextMatrix(i, .ColIndex("NoteIDA")))
                   .TextMatrix(i, .ColIndex("ID")) = (IIf(IsNull(rs.Fields("ID2").value), 0, rs.Fields("ID2").value))
                   .TextMatrix(i, .ColIndex("recordDate")) = (IIf(IsNull(rs.Fields("recordDate").value), 0, rs.Fields("recordDate").value))
                   .TextMatrix(i, .ColIndex("BoxName")) = (IIf(IsNull(rs.Fields("BoxName").value), 0, rs.Fields("BoxName").value))
                   .TextMatrix(i, .ColIndex("account_name")) = (IIf(IsNull(rs.Fields("account_name").value), 0, rs.Fields("account_name").value))
                   .TextMatrix(i, .ColIndex("branch_name")) = (IIf(IsNull(rs.Fields("branch_name").value), 0, rs.Fields("branch_name").value))
                   
                   .TextMatrix(i, .ColIndex("NoteIDA")) = (IIf(IsNull(rs.Fields("notesallid").value), 0, rs.Fields("notesallid").value))
                   
                   .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(rs.Fields("LeaderName").value), "", rs.Fields("LeaderName").value))
                   .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(rs.Fields("ShipID").value), 0, rs.Fields("ShipID").value))
                   .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(rs.Fields("TripNo").value), "", rs.Fields("TripNo").value))
                   .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(rs.Fields("TripDate").value), "", rs.Fields("TripDate").value))
                   .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(rs.Fields("BranchID").value), 0, rs.Fields("BranchID").value))
                   .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(rs.Fields("QtyDownload").value), "", rs.Fields("QtyDownload").value))
                   .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("Price").value), "", rs.Fields("Price").value))
                   .TextMatrix(i, .ColIndex("TotalValue")) = (IIf(IsNull(rs.Fields("TotalValue").value), "", rs.Fields("TotalValue").value))
                   .TextMatrix(i, .ColIndex("Weight")) = (IIf(IsNull(rs.Fields("Weight").value), "", rs.Fields("Weight").value))
                   .TextMatrix(i, .ColIndex("RecNo")) = (IIf(IsNull(rs.Fields("RecNo").value), "", rs.Fields("RecNo").value))
                   .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(rs.Fields("QtyDischarge").value), "", rs.Fields("QtyDischarge").value))
                   .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(rs.Fields("CarType1").value), 1, rs.Fields("CarType1").value))
                   .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(rs.Fields("CardNO").value), "", rs.Fields("CardNO").value))
                    .TextMatrix(i, .ColIndex("BoardNO")) = (IIf(IsNull(rs.Fields("BoardNO").value), "", rs.Fields("BoardNO").value))
                   .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(rs.Fields("CardNO2").value), "", rs.Fields("CardNO2").value))
                   .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(rs.Fields("Remarks").value), "", rs.Fields("Remarks").value))
                   .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(rs.Fields("FromID").value), 0, rs.Fields("FromID").value))
                  .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(rs.Fields("GovernmentName").value), "", rs.Fields("GovernmentName").value))
                    
                    .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), 0, rs.Fields("CusID").value))
                    .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
                  .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(rs.Fields("ToID").value), 0, rs.Fields("ToID").value))
                  .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(rs.Fields("ToGovernmentName").value), "", rs.Fields("ToGovernmentName").value))
                  .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(rs.Fields("CarTypeID").value), 0, rs.Fields("CarTypeID").value))
                  .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(rs.Fields("CarID").value), 0, rs.Fields("CarID").value))
                  If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
                  .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(rs.Fields("BoardNo2").value), "", rs.Fields("BoardNo2").value))
                  Else
                  .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(rs.Fields("BoardNO").value), "", rs.Fields("BoardNO").value))
                  End If
                    .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(rs.Fields("NoteID").value), 0, rs.Fields("NoteID").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(rs.Fields("ShipName").value), "", rs.Fields("ShipName").value))
                    .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value))
                    .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value))
                 Else
                 .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(rs.Fields("ShipNameE").value), "", rs.Fields("ShipNameE").value))
                 .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value))
                 .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(rs.Fields("branch_namee").value), "", rs.Fields("branch_namee").value))
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
  Me.Caption = "Saerch Vehicles Maintenance Plan"

lbprocess.Caption = "Code"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lblLW.Caption = "Saerch By"
lbl(3).Caption = "Car"
 lbl(12).Caption = "Current KM "
lbl(0).Caption = "Select Maint."
Frame1.Caption = "Priod"
lbl(4).Caption = "From"
lbl(2).Caption = "To"


 
    
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
         .TextMatrix(0, .ColIndex("BoardNO")) = "Car"
        .TextMatrix(0, .ColIndex("CurrentKM")) = "CurrentKM"
       .TextMatrix(0, .ColIndex("name")) = "Type Maint"
    End With
  '
End Sub






Private Sub GridInstallments_Click()
If Me.Indx = 3 Then
    
    Nationality.FindRec val(GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("id")))
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = 2))) AND  dbo.TblCarsData.Fullcode like '%" & Text2.Text & "%'  "


   Dcombos.GetEquipments Me.Dccar, str
   
    


End Sub
