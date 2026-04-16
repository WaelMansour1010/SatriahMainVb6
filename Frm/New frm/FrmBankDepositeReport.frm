VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBankDepositeReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmBankDepositeReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10365
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
      _cx             =   1931
      _cy             =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4725
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10395
      Begin VB.TextBox TxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   4125
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1800
         Width           =   3225
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=<"
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
            Index           =   4
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=>"
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
            Index           =   3
            Left            =   -120
            RightToLeft     =   -1  'True
            TabIndex        =   32
            ToolTipText     =   "«ﬂ»— „‰"
            Top             =   0
            Width           =   705
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
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   31
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
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   30
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
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   29
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1800
         Width           =   885
      End
      Begin XtremeSuiteControls.RadioButton RdEda 
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   24
         Top             =   2280
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«Ìœ«⁄ ‰ﬁœÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox txtEmpCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   705
      End
      Begin VB.Frame Frame1 
         Caption         =   "„‰ «·› —Â"
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2760
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   12
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   77529091
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   77529091
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   3
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   4
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   5880
         TabIndex        =   5
         Top             =   120
         Width           =   4455
         Begin VB.Image Image1 
            Height          =   3675
            Left            =   0
            Picture         =   "FrmBankDepositeReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4395
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1095
            Left            =   240
            TabIndex        =   6
            Top             =   3840
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo Dcbank 
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   600
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   1350
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   990
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdEda 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«Ìœ«⁄ ‘Ìﬂ« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdEda 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "‰ﬁœÌ Ê‘Ìﬂ« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”·”· «·«Ìœ«⁄"
         Height          =   285
         Index           =   2
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€"
         Height          =   285
         Index           =   0
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“Ì‰…"
         Height          =   285
         Index           =   14
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   15
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   240
         Index           =   64
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "‘«‘…  ﬁ«—Ì— «·«Ìœ«⁄«  «·»‰ﬂÌÂ "
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
         Height          =   1020
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3480
         Width           =   5535
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Top             =   3480
         Width           =   5655
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5640
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   9
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·«Ìœ«⁄«  «·»‰ﬂÌÂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   10305
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmBankDepositeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Private Sub btnClear_Click()
clear_all Me

DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

GetData
          
        Case 1
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub




Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
  
  
    Set Dcombos = New ClsDataCombos
 Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetBanks Me.Dcbank
    Dcombos.GetBoxes Me.DcboBox

DtpDateFrom.value = ""
DtpDateTo.value = ""
    Resize_Form Me

End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then txtEmpCode.text = "": Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    txtEmpCode.text = EmpCode
    
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
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

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblBanksDepositeDetails.TblBanksDepositeId, dbo.TblBanksDepositeDetails.box_or_bank, dbo.TblBanksDepositeDetails.[value], "
StrSQL = StrSQL & "                      dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.Remarks, dbo.TblBanksDepositeDetails.BoxID, dbo.TblBoxesData.BoxName,"
StrSQL = StrSQL & "                      dbo.TblBoxesData.BoxNameE, dbo.TblBanksDepositeDetails.BankName AS DetBankName, dbo.TblBanksDepositeDetails.DueDate,"
StrSQL = StrSQL & "                      dbo.TblBanksDeposite.NoteSerial1, dbo.TblBanksDeposite.NoteSerial, dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.bankid,"
StrSQL = StrSQL & "                      BanksData_1.BankName AS DepositeBankName, dbo.TblBanksDeposite.id, dbo.TblBanksDeposite.Remarks AS Remarkss, dbo.TblBanksDeposite.Emp_id,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblBanksDeposite.OldNoteSerial1, BanksData_1.BankNamee AS DepositeBankNameE, BanksData_2.BankName,"
StrSQL = StrSQL & "                      BanksData_2.BankNamee"
StrSQL = StrSQL & " FROM         dbo.TblBanksDepositeDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.BanksData BanksData_1 ON dbo.TblBanksDepositeDetails.bankid = BanksData_1.BankID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBanksDeposite ON dbo.TblBanksDepositeDetails.TblBanksDepositeId = dbo.TblBanksDeposite.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblBanksDeposite.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.BanksData BanksData_2 ON dbo.TblBanksDeposite.bankid = BanksData_2.BankID"
StrSQL = StrSQL & " where 1=1"
If Me.Dcbank.text <> "" And val(Dcbank.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDeposite.bankid = " & val(Me.Dcbank.BoundText)

End If
If Me.DcboEmpName.text <> "" And val(DcboEmpName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblEmployee.Emp_ID = " & val(Me.DcboEmpName.BoundText)

End If
If Me.DcboBox.text <> "" And val(DcboBox.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblBoxesData.BoxID = " & val(Me.DcboBox.BoundText)

End If
If Me.TxtNoteSerial1.text <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDeposite.NoteSerial1 ='" & TxtNoteSerial1.text & "'"

End If
If Me.RdEda(0).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.box_or_bank =0"

End If
If Me.RdEda(1).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.box_or_bank =1"

End If
If Me.TxtValue.text <> "" Then
If Opt(2).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.value >" & TxtValue.text & ""
 ElseIf Opt(1).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.value =" & TxtValue.text & ""
  ElseIf Opt(0).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.value <" & TxtValue.text & ""
   ElseIf Opt(4).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.value >=" & TxtValue.text & ""
  ElseIf Opt(3).value = True Then
    StrSQL = StrSQL & " AND   dbo.TblBanksDepositeDetails.value <=" & TxtValue.text & ""
    
  End If
End If

 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBanksDeposite.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBanksDeposite.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBankDepositeReport.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBankDepositeReport.rpt"
            
       End If



    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value

    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If

  Dim Total As String
  Dim totl As Double


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function



 
Private Sub TxtValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub
