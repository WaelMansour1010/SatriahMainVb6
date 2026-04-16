VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmGeneralFundReceiptSerch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   Icon            =   "FrmGeneralFundReceiptSerch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   13455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ ⁄‰ ”‰œ ﬁ»÷ «·’‰œÊﬁ «·⁄«„  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   5400
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmGeneralFundReceiptSerch.frx":6852
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   25
      Top             =   720
      Width           =   13455
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2625
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   13155
         _cx             =   23204
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmGeneralFundReceiptSerch.frx":15141
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   6120
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   285
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   6720
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   18
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
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
         ButtonImage     =   "FrmGeneralFundReceiptSerch.frx":152A9
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
         Left            =   5160
         TabIndex        =   19
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
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
         ButtonImage     =   "FrmGeneralFundReceiptSerch.frx":1BB0B
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
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
         ButtonImage     =   "FrmGeneralFundReceiptSerch.frx":2236D
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
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3720
      Width           =   5955
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   195
         Index           =   14
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   7335
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84672515
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   84672515
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   195
         Index           =   13
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   4
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   3
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   6855
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   3120
            TabIndex        =   38
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   84672515
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   330
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   84672515
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   195
            Index           =   12
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   315
            Index           =   11
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   10
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   120
         Width           =   6855
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   3120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   84672515
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   84672515
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   9
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   315
            Index           =   8
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   195
            Index           =   1
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   825
         End
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   6960
         TabIndex        =   2
         Top             =   720
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   6960
         TabIndex        =   29
         Top             =   1200
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„‰œÊ»"
         Height          =   285
         Index           =   0
         Left            =   11160
         TabIndex        =   30
         Top             =   1200
         Width           =   2205
      End
      Begin VB.Label lblLL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Left            =   11160
         TabIndex        =   4
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰œÊﬁ"
         Height          =   285
         Index           =   7
         Left            =   11160
         TabIndex        =   3
         Top             =   720
         Width           =   2205
      End
   End
End
Attribute VB_Name = "FrmGeneralFundReceiptSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As FrmGeneralFundReceipt
Private Sub Fg_Click()
FrmGeneralFundReceipt.FindRec val(Fg.TextMatrix(Fg.Row, 1))
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
     Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetSalesRepData Me.DataCombo1
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
     SetDtpickerDate Me.DTPicker1
     SetDtpickerDate Me.DTPicker2
     SetDtpickerDate Me.DTPicker3
     SetDtpickerDate Me.DTPicker4
   End Sub
   Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
          Me.DTPicker1.value = ""
          Me.DTPicker2.value = ""
          Me.DTPicker3.value = ""
          Me.DTPicker4.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
            Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT      dbo.TBLGeneralFundReceipt.IDGFR, dbo.TBLGeneralFundReceipt.ManualNo, dbo.TBLGeneralFundReceipt.DateM, dbo.TBLGeneralFundReceipt.DateH, dbo.TBLGeneralFundReceipt.BranchID,"
    sql = sql + "      dbo.TBLGeneralFundReceipt.GeneralBoxID, dbo.TBLGeneralFundReceipt.DelegateID, dbo.TBLGeneralFundReceipt.FromDate, dbo.TBLGeneralFundReceipt.ToDate,"
    sql = sql + "      dbo.TBLGeneralFundReceipt.Explan, dbo.TBLGeneralFundReceipt.TotallVal, dbo.TBLGeneralFundReceipt.UserID, TblBranchesData_1.branch_id, TblBranchesData_1.branch_name,"
    sql = sql + "      TblBranchesData_1.branch_namee, TblBoxesData_1.BoxID, TblBoxesData_1.BoxName, TblBoxesData_1.BoxNameE, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
    sql = sql + "      dbo.TblEmployee.Emp_Namee"
    sql = sql + "      FROM         dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
    sql = sql + "      dbo.TblEmployee RIGHT OUTER JOIN"
    sql = sql + "      dbo.TBLGeneralFundReceipt ON dbo.TblEmployee.Emp_ID = dbo.TBLGeneralFundReceipt.DelegateID LEFT OUTER JOIN"
    sql = sql + "     dbo.TblBoxesData TblBoxesData_1 ON dbo.TBLGeneralFundReceipt.GeneralBoxID = TblBoxesData_1.BoxID ON TblBranchesData_1.branch_id = dbo.TBLGeneralFundReceipt.BranchID"
    
       BolBegine = False
       StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TBLGeneralFundReceipt.IDGFR >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLGeneralFundReceipt.IDGFR >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.IDGFR <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TBLGeneralFundReceipt.IDGFR <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLGeneralFundReceipt.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLGeneralFundReceipt.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBranchesData_1.branch_id =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblBranchesData_1.branch_id =" & Me.Dcbranch.BoundText & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Me.DcboBox.text <> "" And (val(DcboBox.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND TblBoxesData_1.BoxID =" & Me.DcboBox.BoundText & ""
        Else
          BolBegine = True
          StrWhere = " Where TblBoxesData_1.BoxID =" & Me.DcboBox.BoundText & ""
       End If
      End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If Me.DataCombo1.text <> "" And (val(DataCombo1.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.DelegateID =" & Me.DataCombo1.BoundText & ""
        Else
          BolBegine = True
          StrWhere = " Where dbo.TBLGeneralFundReceipt.DelegateID =" & Me.DataCombo1.BoundText & ""
       End If
      End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.FromDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLGeneralFundReceipt.FromDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.FromDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  ddbo.TBLGeneralFundReceipt.FromDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         If Not IsNull(Me.DTPicker3.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.ToDate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLGeneralFundReceipt.ToDate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DTPicker4.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLGeneralFundReceipt.ToDate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLGeneralFundReceipt.ToDate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TBLGeneralFundReceipt.IDGFR"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDGFR").value), "", rs("IDGFR").value)
                If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("DateM").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("StatmentDT")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
               .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               .TextMatrix(i, .ColIndex("StatmentDT")) = IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
               .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
               .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
               .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
               .TextMatrix(i, .ColIndex("Explan")) = IIf(IsNull(rs("Explan").value), "", rs("Explan").value)
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataAdvince()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
     
    sql = "SELECT   dbo.TBLBankSettlement.IDBS, dbo.TBLBankSettlement.DateM, dbo.TBLBankSettlement.DateH, dbo.TBLBankSettlement.BranchID, dbo.TBLBankSettlement.SettlementDT,"
    sql = sql & "     dbo.TBLBankSettlement.BankID, dbo.TBLBankSettlement.FromDT, dbo.TBLBankSettlement.ToDT, dbo.TBLBankSettlement.EXPCheck, dbo.TBLBankSettlement.UserID,"
    sql = sql & "    dbo.TBLBankSettlementJoin.ID, dbo.TBLBankSettlementJoin.IDBS AS IDBSJON, dbo.TBLBankSettlementJoin.MoveDT, dbo.TBLBankSettlementJoin.BankValue, dbo.TBLBankSettlementJoin.BankRF,"
    sql = sql & "     dbo.TBLBankSettlementJoin.Explan , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE, dbo.BanksData.BankName, dbo.BanksData.BankNamee"
    sql = sql & "     FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
    sql = sql & "   dbo.TBLBankSettlement LEFT OUTER JOIN"
    sql = sql & "     dbo.BanksData ON dbo.TBLBankSettlement.BankID = dbo.BanksData.BankID ON dbo.TblBranchesData.branch_id = dbo.TBLBankSettlement.BranchID LEFT OUTER JOIN"
    sql = sql & "     dbo.TBLBankSettlementJoin ON dbo.TBLBankSettlement.IDBS = dbo.TBLBankSettlementJoin.IDBS"
  '  sql = sql & "      dbo.TBLBankSettlement.IDBS = dbo.TBLBankSettlementJoin.IDBS"
    
       BolBegine = False
       StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TBLBankSettlement.IDBS >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLBankSettlement.IDBS >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TBLBankSettlement.IDBS <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TBLBankSettlement.IDBS <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLBankSettlement.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLBankSettlement.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLBankSettlement.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLBankSettlement.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TBLBankSettlement.BranchID =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TBLBankSettlement.BranchID =" & Me.Dcbranch.BoundText & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Me.DcboBox.text <> "" And (val(DcboBox.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TBLBankSettlement.BankID =" & Me.DcboBox.BoundText & ""
        Else
          BolBegine = True
          StrWhere = " Where dbo.TBLBankSettlement.BankID =" & Me.DcboBox.BoundText & ""
       End If
      End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '     If (Me.Check1.value = True) Then
   '                 If BolBegine = True Then
   '                 StrWhere = StrWhere & " AND  dbo.TBLBankSettlement.EXPCheck = 1 "
   '                  Else
   '                 BolBegine = True
   '                 StrWhere = StrWhere & " Where  dbo.TBLBankSettlement.EXPCheck = 0 "
   '                 End If
   '     End If
   '
'       If (Me.Check2.value = True) Then
'        If BolBegine = True Then
'        StrWhere = StrWhere & " AND  dbo.TBLBankSettlement.EXPCheck = 0 "
'        Else
'        BolBegine = True
'        StrWhere = StrWhere & " Where  dbo.TBLBankSettlement.EXPCheck = 1 "
'        End If
'        End If
    '-----------------------------------
    ' advinse search
   If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLBankSettlementJoin.MoveDT >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLBankSettlementJoin.MoveDT >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLBankSettlementJoin.MoveDT <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLBankSettlementJoin.MoveDT <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' If Me.Text1.text <> "" Then
       ' If BolBegine = True Then
       '     StrWhere = StrWhere & " AND  dbo.TBLBankSettlementJoin.BankValue like '%" & Me.Text1.text & "%'"
       ' Else
       '     BolBegine = True
       '     StrWhere = " Where  dbo.TBLBankSettlementJoin.BankValue like '%" & Me.Text1.text & "%'"
       ' End If
       'End If
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '      If Me.oldTxtSerial1.text <> "" Then
  '      If BolBegine = True Then
  '          StrWhere = StrWhere & " AND  dbo.TBLBankSettlementJoin.BankRF like '%" & Me.oldTxtSerial1.text & "%'"
  '      Else
  '          BolBegine = True
  '          StrWhere = " Where  dbo.TBLBankSettlementJoin.BankRF like '%" & Me.oldTxtSerial1.text & "%'"
        'End If
       'End If
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     '   If Me.txtto.text <> "" Then
     '   If BolBegine = True Then
     '       StrWhere = StrWhere & " AND dbo.TBLBankSettlementJoin.Explan like '%" & Me.txtto.text & "%'"
     '   Else
     '       BolBegine = True
   '         StrWhere = " Where  dbo.TBLBankSettlementJoin.Explan like '%" & Me.txtto.text & "%'"
   '     End If
   '    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TBLBankSettlement.IDBS"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
            With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDBS").value), "", rs("IDBS").value)
                If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("DateM").value, "yyyy/M/d")
                End If
                If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("StatmentDT")) = Format(rs("SettlementDT").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
                Else
               .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_nameE").value), "", rs("branch_nameE").value)
               .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankNamee").value), "", rs("BankNamee").value)
                End If
               .TextMatrix(i, .ColIndex("StatType")) = IIf(IsNull(rs("EXPCheck").value), "", rs("EXPCheck").value)
               '''''''''''''''''''''''''''''''''''''''''''''''
                If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("MovDT")) = Format(rs("MoveDT").value, "yyyy/M/d")
                End If
               .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs("BankValue").value), "", rs("BankValue").value)
               .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(rs("BankRF").value), "", rs("BankRF").value)
               .TextMatrix(i, .ColIndex("Explan")) = IIf(IsNull(rs("Explan").value), "", rs("Explan").value)
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
     Me.Caption = "Bank Settlement Search"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Operation ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Operation Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    Me.lblLL.Caption = "Branch"
    Me.lbl(7).Caption = "Bank Name"
   ' Me.Check1.Caption = "Manual"
   ' Me.Check2.Caption = "Import File"
    'Me.CmdShowMoreOptions.Caption = "Advanced"
    'Me.Option1.Caption = "All"
    lbl(2).Caption = "Total"
    'Me.ISButton2.Caption = "Advanced"
    ''''''''''''''''''''''' next
    Me.lbl(0).Caption = "Movement Date"
    Me.lbl(1).Caption = "From"
    Me.lbl(8).Caption = "To"
    Me.Label1(9).Caption = "Value"
    'Me.Labelbnkrf.Caption = "Bank Reference"
    'Me.lbl(9).Caption = "Explanation"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("StatmentDT")) = "Settlement Date"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("StatType")) = "Settlement Type"
        .TextMatrix(0, .ColIndex("MovDT")) = "Movement Date"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("MovDT")) = "Movement Date"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("BankRF")) = "Bank Reference"
        .TextMatrix(0, .ColIndex("Explan")) = "Value"
        .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
    End With
  End Sub
'''''''''''''''''''''''''''' end


