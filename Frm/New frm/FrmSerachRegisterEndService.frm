VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSerachRegisterEndService 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ »Ì«‰«   —ﬂ Œœ„… "
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "FrmSerachRegisterEndService.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì… «·⁄ﬁœ"
      Height          =   555
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3240
      Width           =   5415
      Begin MSComCtl2.DTPicker DateEndContracFrom 
         Height          =   330
         Left            =   2160
         TabIndex        =   31
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DateEndContracTo 
         Height          =   330
         Left            =   90
         TabIndex        =   32
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   15
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   14
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·«” ﬁ«·…"
      Height          =   555
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   5415
      Begin MSComCtl2.DTPicker DTP_DateFrom 
         Height          =   330
         Left            =   2160
         TabIndex        =   26
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DTP_DateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   27
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   13
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   12
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3840
      Width           =   11025
      Begin VB.ComboBox dcjopstatus12 
         Height          =   315
         ItemData        =   "FrmSerachRegisterEndService.frx":038A
         Left            =   5880
         List            =   "FrmSerachRegisterEndService.frx":038C
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox dcjopstatus1 
         Height          =   315
         ItemData        =   "FrmSerachRegisterEndService.frx":038E
         Left            =   5760
         List            =   "FrmSerachRegisterEndService.frx":0390
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox TxtTelephone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   915
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   360
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   5760
         TabIndex        =   21
         Top             =   720
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboJobsType 
         Height          =   315
         Left            =   5760
         TabIndex        =   23
         Top             =   1080
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbMangment 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDirctManger 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcNational 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcjopstatus 
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Tag             =   "Õœœ «·Õ«·…"
         Top             =   1440
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ«·Â"
         Height          =   285
         Index           =   4
         Left            =   4785
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«‘⁄«—"
         Height          =   285
         Index           =   5
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰"
         Height          =   285
         Index           =   17
         Left            =   9600
         TabIndex        =   42
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   285
         Index           =   16
         Left            =   4650
         TabIndex        =   40
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ— «·„»«‘—"
         Height          =   285
         Index           =   8
         Left            =   4440
         TabIndex        =   38
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„"
         Height          =   285
         Index           =   11
         Left            =   9600
         TabIndex        =   36
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›Â"
         Height          =   285
         Index           =   9
         Left            =   9600
         TabIndex        =   24
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—Â"
         Height          =   285
         Index           =   7
         Left            =   4440
         TabIndex        =   22
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   0
         Left            =   9750
         TabIndex        =   20
         Top             =   375
         Width           =   1125
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·ÿ·»"
      Height          =   555
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   4575
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2040
         TabIndex        =   6
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   111411203
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·ÌÂ"
      Height          =   645
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
      Width           =   3555
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2775
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
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10995
      _cx             =   19394
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSerachRegisterEndService.frx":0392
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
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   12
      Top             =   6000
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
      TabIndex        =   13
      Top             =   6000
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
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSerachRegisterEndService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Ind As Integer
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 
 GetData
           
        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
Me.DTP_DateFrom.value = ""
Me.DTP_DateTo.value = ""
Me.DateEndContracFrom.value = ""
Me.DateEndContracTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub



Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    End Sub

Private Sub Fg_Click()
If Ind = 1 Then
End_oF_service.TxtReqNo.text = (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
End_oF_service.TxtReqNo_KeyPress (13)
Else
FrmRegisterHoliday.FindRec (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
End If
End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  Set Dcombos = New ClsDataCombos
    
    
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmpJobsTypes Me.DcboJobsType
    Dcombos.GetEmpDepartments Me.DcbDept
    '''//
    
    Dcombos.GETNationality Me.DcNational
    Dcombos.GetSection Me.DcbMangment
    Dcombos.GetEmployees Me.DcbDirctManger
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select  id,name  from jopstatus   where id>1"
    Else
        My_SQL = "  select  id,namee  from jopstatus where id>1  "
    End If
    fill_combo dcjopstatus, My_SQL
If SystemOptions.UserInterface = ArabicInterface Then
dcjopstatus1.AddItem "«” ﬁ«·…"
dcjopstatus1.AddItem "⁄œ„ «·—€»… ›Ì «· ÃœÌœ"
dcjopstatus1.AddItem "«”»«» „—÷Ì…"
dcjopstatus1.AddItem "«Œ—Ï"

dcjopstatus12.AddItem "«” ﬁ«·…"
dcjopstatus12.AddItem "⁄œ„ «·—€»… ›Ì «· ÃœÌœ"
dcjopstatus12.AddItem "«”»«» „—÷Ì…"
dcjopstatus12.AddItem "«Œ—Ï"
Else
dcjopstatus1.AddItem "Resigantion"
dcjopstatus1.AddItem "Non-Renewal of Contract"
dcjopstatus1.AddItem "Sick Leave"
dcjopstatus1.AddItem "Other"

dcjopstatus12.AddItem "Resigantion"
dcjopstatus12.AddItem "Non-Renewal of Contract"
dcjopstatus12.AddItem "Sick Leave"
dcjopstatus12.AddItem "Other"

End If
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
     SetDtpickerDate Me.DTP_DateFrom
    SetDtpickerDate Me.DTP_DateTo
     SetDtpickerDate Me.DateEndContracFrom
    SetDtpickerDate Me.DateEndContracTo

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

StrSQL = "SELECT     dbo.TBLRegisterHoliday.id, dbo.TBLRegisterHoliday.EmpID, dbo.TBLRegisterHoliday.EndWork, dbo.TBLRegisterHoliday.Notsstkala, "
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.jopstatusid, dbo.jopstatus.name, dbo.jopstatus.namee, dbo.TBLRegisterHoliday.BranchID, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TBLRegisterHoliday.NationID, dbo.Nationality.name AS Nationaliname, dbo.Nationality.namee AS NationalinameE,"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TBLRegisterHoliday.DirctMangerID,"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.DepartMentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.MangmentID, dbo.TblSection.name AS Mangname, dbo.TblSection.namee AS MangnameE, dbo.TBLRegisterHoliday.Telephone,"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.LogConract, dbo.TBLRegisterHoliday.Other, dbo.TBLRegisterHoliday.Jopstatus1, dbo.TBLRegisterHoliday.DateSatrContrac,"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday.DateEndContrac, dbo.TBLRegisterHoliday.Remarkss, dbo.TBLRegisterHoliday.Des, dbo.TBLRegisterHoliday.RecordDate,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Code, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name4, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee3, TblEmployee_1.Emp_Namee4, TblEmployee_2.Emp_Name AS MangerEmp_Name,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Namee AS MangerEmp_NameE, dbo.TBLRegisterHoliday.Salary"
StrSQL = StrSQL & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLRegisterHoliday ON TblEmployee_1.Emp_ID = dbo.TBLRegisterHoliday.EmpID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TBLRegisterHoliday.DirctMangerID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblSection ON dbo.TBLRegisterHoliday.MangmentID = dbo.TblSection.Id ON"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TBLRegisterHoliday.DepartMentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.jopstatus ON dbo.TBLRegisterHoliday.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TBLRegisterHoliday.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Nationality ON dbo.TBLRegisterHoliday.NationID = dbo.Nationality.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TBLRegisterHoliday.BranchID = dbo.TblBranchesData.branch_id"
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TBLRegisterHoliday.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
        If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLRegisterHoliday.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLRegisterHoliday.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    ''''///
       If Not IsNull(Me.DateEndContracFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.DateEndContrac >=" & SQLDate(Me.DateEndContracFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.DateEndContrac >=" & SQLDate(Me.DateEndContracFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DateEndContracTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLRegisterHoliday.DateEndContrac <=" & SQLDate(Me.DateEndContracTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLRegisterHoliday.DateEndContrac <=" & SQLDate(Me.DateEndContracTo.value, True) & ""
        End If
    End If
    ''/////
        If Not IsNull(Me.DTP_DateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.EndWork >=" & SQLDate(Me.DTP_DateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.EndWork >=" & SQLDate(Me.DTP_DateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTP_DateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLRegisterHoliday.EndWork <=" & SQLDate(Me.DTP_DateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLRegisterHoliday.EndWork <=" & SQLDate(Me.DTP_DateTo.value, True) & ""
        End If
    End If
    
      If Me.TxtSearchCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  TblEmployee_1.fullcode ='" & Me.TxtSearchCode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where  TblEmployee_1.fullcode ='" & Me.TxtSearchCode.text & "'"
        End If
    End If
    If Me.DcboEmpName.text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.EmpID =" & Me.DcboEmpName.BoundText & ""
        End If
    End If
        If Me.DcbDirctManger.text <> "" And (val(DcbDirctManger.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.DirctMangerID =" & Me.DcbDirctManger.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.DirctMangerID =" & Me.DcbDirctManger.BoundText & ""
        End If
    End If
    
    
        If Me.DcbMangment.text <> "" And (val(DcbMangment.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.MangmentID =" & Me.DcbMangment.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.MangmentID =" & Me.DcbMangment.BoundText & ""
        End If
    End If
       If Me.DcbDept.text <> "" And (val(DcbDept.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.DepartMentID =" & Me.DcbDept.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.DepartMentID =" & Me.DcbDept.BoundText & ""
        End If
    End If
    
      If Me.DcbDirctManger.text <> "" And (val(DcbDirctManger.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.DirctMangerID =" & Me.DcbDirctManger.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.DirctMangerID =" & Me.DcbDirctManger.BoundText & ""
        End If
    End If
       If Me.DcboJobsType.text <> "" And (val(DcboJobsType.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.JobID =" & Me.DcboJobsType.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.JobID =" & Me.DcboJobsType.BoundText & ""
        End If
    End If
        If Me.DcNational.text <> "" And (val(DcNational.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.NationID =" & Me.DcNational.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.NationID =" & Me.DcNational.BoundText & ""
        End If
    End If
    
         If Me.TxtTelephone.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.Telephone ='" & Me.TxtTelephone.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.Telephone ='" & Me.TxtTelephone.text & "'"
        End If
    End If
    
   
       If Me.dcjopstatus1.text <> "" And (val(dcjopstatus1.ListIndex) <> -1) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.Jopstatus1 =" & val(Me.dcjopstatus1.ListIndex) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.Jopstatus1 =" & val(Me.dcjopstatus1.ListIndex) & ""
        End If
    End If
    
          If Me.dcjopstatus.text <> "" And (val(dcjopstatus.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLRegisterHoliday.jopstatusid =" & Me.dcjopstatus.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLRegisterHoliday.jopstatusid =" & Me.dcjopstatus.BoundText & ""
        End If
    End If
       



    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TBLRegisterHoliday.ID"
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
               
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                .TextMatrix(i, .ColIndex("LogConract")) = IIf(IsNull(rs("LogConract").value), "", rs("LogConract").value)
                 If Not (IsNull(rs("DateEndContrac").value)) Then
                    .TextMatrix(i, .ColIndex("DateEndContrac")) = Format(rs("DateEndContrac").value, "yyyy/M/d")
                End If
                
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("EndWork").value)) Then
                    .TextMatrix(i, .ColIndex("EndWork")) = Format(rs("EndWork").value, "yyyy/M/d")
                End If
                dcjopstatus12.ListIndex = IIf(IsNull(rs("Jopstatus1").value), -1, rs("Jopstatus1").value)
                .TextMatrix(i, .ColIndex("Jopstatus1")) = dcjopstatus12.text
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
            .TextMatrix(i, .ColIndex("Nationaliname")) = IIf(IsNull(rs("Nationaliname").value), "", rs("Nationaliname").value)
            .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
            .TextMatrix(i, .ColIndex("MangerEmp_Name")) = IIf(IsNull(rs("MangerEmp_Name").value), "", rs("MangerEmp_Name").value)
            .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
            .TextMatrix(i, .ColIndex("Mangname")) = IIf(IsNull(rs("Mangname").value), "", rs("Mangname").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
          
            Else
            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            .TextMatrix(i, .ColIndex("Nationaliname")) = IIf(IsNull(rs("NationalinameE").value), "", rs("NationalinameE").value)
            .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
            .TextMatrix(i, .ColIndex("MangerEmp_Name")) = IIf(IsNull(rs("MangerEmp_NameE").value), "", rs("MangerEmp_NameE").value)
            .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
            .TextMatrix(i, .ColIndex("Mangname")) = IIf(IsNull(rs("MangnameE").value), "", rs("MangnameE").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
            End If
          ' .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
          ' .TextMatrix(i, .ColIndex("NumEkama")) = IIf(IsNull(rs("NumEkama").value), "", rs("NumEkama").value)
         '  If SystemOptions.UserInterface = ArabicInterface Then
         '      If rs("SpecificHolidyaType1").value = True Then
         '      .TextMatrix(i, .ColIndex("typevocation")) = "«÷ÿ—«—ÌÂ"
         '      Else
         '        .TextMatrix(i, .ColIndex("typevocation")) = "—”„ÌÂ"
         '      End If
         '      Else
         '       If rs("SpecificHolidyaType1").value = True Then
         '      .TextMatrix(i, .ColIndex("typevocation")) = "Forced"
         '      Else
         '        .TextMatrix(i, .ColIndex("typevocation")) = "Official"
         '      End If
         '      End If

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
  Me.Caption = "Saerch Data of Register End Work"
lbprocess.Caption = "No.Req"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbreg.Caption = "Date.Req"
Frame1.Caption = "Date Resigantion"
lbl(13).Caption = "From"
lbl(12).Caption = "To"
Frame2.Caption = "Date End Work"
lbl(14).Caption = "From"
lbl(15).Caption = "To"
lblLW.Caption = "Saerch By"
lbl(2).Caption = "Total"

lbl(0).Caption = "Employee"
lbl(11).Caption = "Department"
lbl(7).Caption = "Management "
lbl(8).Caption = "Manager"
lbl(9).Caption = "Job"
lbl(16).Caption = "Nationality"
lbl(17).Caption = "Telephone"
Label1(4).Caption = "Status"
Label1(5).Caption = "Notice"





     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No.Req"
         .TextMatrix(0, .ColIndex("RecordDate")) = "Date.Req"
         .TextMatrix(0, .ColIndex("EndWork")) = "Date.Resigantion"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
       ' .TextMatrix(0, .ColIndex("GroupName")) = "Location"
         .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
       .TextMatrix(0, .ColIndex("Mangname")) = "Management"
       .TextMatrix(0, .ColIndex("DepartmentName")) = "Department"
       
         .TextMatrix(0, .ColIndex("MangerEmp_Name")) = "Manger"
         .TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        .TextMatrix(0, .ColIndex("Nationaliname")) = "Nationality"
        .TextMatrix(0, .ColIndex("DateSatrContrac")) = "Bign Work"
         .TextMatrix(0, .ColIndex("LogConract")) = "LogConract"
       .TextMatrix(0, .ColIndex("name")) = "Status"
       .TextMatrix(0, .ColIndex("Jopstatus1")) = "Notice"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
End Sub
