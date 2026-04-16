VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmOrboon 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmOrboon.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10425
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
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93782017
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93782017
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4605
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   10395
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «·⁄—»Ê‰"
         Height          =   600
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1920
         Width           =   6555
         Begin XtremeSuiteControls.CheckBox ChkOrboon 
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   36
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄—»Ê‰ „”œœ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOrboon 
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   37
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄—»Ê‰ „·€Ì"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkOrboon 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "⁄—»Ê‰ „— Ã⁄"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox TxtEmployeeID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9720
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2520
         Width           =   6555
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   20
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   93782017
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   21
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   93782017
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   3
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   6960
         TabIndex        =   16
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmOrboon.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3300
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
            Height          =   5295
            Left            =   120
            TabIndex        =   17
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   14
         Top             =   4680
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   1560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   1200
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   5640
         TabIndex        =   34
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
         Top             =   4560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·„‰œÊ» „Õœœ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   4560
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·ÊÕœ… „ÕœœÂ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÿ»ﬁ« ·‰Ê⁄ „Õœœ"
         Height          =   195
         Index           =   15
         Left            =   5505
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   855
         Left            =   0
         Top             =   3720
         Width           =   6975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
         Height          =   690
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3840
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5550
         TabIndex        =   8
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   0
         Left            =   5595
         TabIndex        =   5
         Top             =   480
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5400
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
      Top             =   5400
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   600
      Picture         =   "FrmOrboon.frx":10A48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·⁄—»Ê‰"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10455
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
Attribute VB_Name = "FrmOrboon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim MSGType As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
'

Private Sub btnClear_Click()
Cmd_Click (1)
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
   
            clear_all Me
        '    ChekCommission.value = vbUnchecked
         Fromdate.value = ""
    ToDate.value = ""
  ChkOrboon(0).value = vbUnchecked
      ChkOrboon(1).value = vbUnchecked
        ChkOrboon(2).value = vbUnchecked
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






Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
DcbUnitType_Change
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , dcbAqarType.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
End Sub

'Private Sub dcsupplier_Change()
'    dcsupplier_Click (0)
'End Sub

'Private Sub dcsupplier_Click(Area As Integer)
'  If val(dcsupplier.BoundText) = 0 Then Exit Sub
'
'    Dim EmpCode  As String
'
   ' GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
'    Me.txtCodeOwner.text = EmpCode
'End Sub

Private Sub DcboEmp_Change()
 If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(dcbAqarType.BoundText) > 0 Then
idd = val(dcbAqarType.BoundText)

idd1 = val(DcbUnitType.BoundText)

Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"

'Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo

End If
End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
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
    
    Dcombos.GetIqar dcbAqarType
    
   ' Dcombos.GetAlarm Me.DcbAlarm
  Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
  '  Dcombos.GetCustomersSuppliers 1, Me.dcCustomer
    
   ' Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
    
    'Dcombos.getAkarUnit Me.DCAkarUnit
    
   ' Dcombos.GetSalesRepData Me.dcbSalesSpec
    
   ' Dcombos.GetCustomersSuppliers 1, Me.dbcClient
    
    Dcombos.GetBranches DcbBranch
    
 ' Dcombos.GetRentStatus dbcAqarStatus
    ChkOrboon(0).value = vbUnchecked
      ChkOrboon(1).value = vbUnchecked
        ChkOrboon(2).value = vbUnchecked
    Fromdate.value = ""
    ToDate.value = ""
    Cmd_Click (1)
    
    Set cSearch = New clsDCboSearch
  '  My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
   ' RsSavRec.CursorLocation = adUseClient
   ' RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
 
    
    Resize_Form Me
    
   
    
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
   
    'gr = 9
    'Order = 9
 MSGType = MsgBox("Â·  —€» ›Ì ≈ŸÂ«— «·⁄—»Ê‰ «·„— »ÿ »«·⁄ﬁÊœ «„ ·«  ", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblAqrEarnest.ID, dbo.TblAqrEarnest.CoustomerName, dbo.TblAqrEarnest.Telephone, dbo.TblAqrEarnest.RecordDate, "
StrSQL = StrSQL & "                      dbo.TblAqrEarnest.RecordDateH, dbo.TblAqar.aqarname, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqrEarnest.UnitNo,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblAqrEarnest.Earnest, dbo.TblAqrEarnest.ValidityDate, dbo.TblAqrEarnest.ValidityDateH,"
StrSQL = StrSQL & "                      dbo.TblAqrEarnest.StatusEarnest, dbo.TblAqrEarnest.NoteID, dbo.TblAqarDetai.unittype, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.branch_no,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee, dbo.Notes.Rent, dbo.Notes.Water, dbo.Notes.commission, dbo.Notes.Instrunce"
StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqrEarnest ON dbo.Notes.NoteID = dbo.TblAqrEarnest.NoteID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblAkarUnit.id = dbo.TblAqarDetai.unittype LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid ON dbo.TblAqrEarnest.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " WHERE     (1 = 1)  "
Else
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblAqrEarnest.ID, dbo.TblAqrEarnest.CoustomerName, dbo.TblAqrEarnest.Telephone, dbo.TblAqrEarnest.RecordDate, "
StrSQL = StrSQL & "                      dbo.TblAqrEarnest.RecordDateH, dbo.TblAqar.aqarname, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqrEarnest.UnitNo,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.unitno AS unitnoName, dbo.TblAqrEarnest.Earnest, dbo.TblAqrEarnest.ValidityDate, dbo.TblAqrEarnest.ValidityDateH,"
StrSQL = StrSQL & "                      dbo.TblAqrEarnest.StatusEarnest, dbo.TblAqrEarnest.NoteID, dbo.TblAqarDetai.unittype, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.branch_no,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.commission, dbo.Notes.Instrunce,"
StrSQL = StrSQL & "                      dbo.GetOrbon(dbo.Notes.NoteSerial1) AS serial"
StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqrEarnest ON dbo.Notes.NoteID = dbo.TblAqrEarnest.NoteID LEFT OUTER JOIN"
StrSQL = StrSQL & "                     dbo.TblAkarUnit RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblAkarUnit.id = dbo.TblAqarDetai.unittype LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblAqarDetai.Aqarid = dbo.TblAqar.Aqarid ON dbo.TblAqrEarnest.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  Where (1 = 1) And (dbo.GetOrbon(dbo.Notes.NoteSerial1) = 0)"
  End If
    BolBegine = False
    StrWhere = ""
    ' dbo.GetOrbon(dbo.Notes.NoteSerial1) AS serial
    
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.Notes.branch_no = " & val(Me.DcbBranch.BoundText)

End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.Text <> "" Then

StrWhere = StrWhere & " AND dbo.TblAqarDetai.Aqarid = " & val(Me.dcbAqarType.BoundText)

End If


'If val(Me.DcboEmp.BoundText) <> 0 Or Me.DcboEmp.text <> "" Then
'StrWhere = StrWhere & " AND dbo.TblAqarCommissions.EmpID  = " & val(DcboEmp.BoundText)

'End If

If val(Me.DcbUnitType.BoundText) <> 0 Or Me.DcbUnitType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAkarUnit.id  = " & val(DcbUnitType.BoundText)

End If
If val(Me.DcbUnitNo.BoundText) <> 0 Or Me.DcbUnitNo.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqrEarnest.UnitNo  = " & val(DcbUnitNo.BoundText)

End If
If Me.ChkOrboon(0).value = vbChecked And Me.ChkOrboon(1).value = vbChecked And Me.ChkOrboon(2).value = vbChecked Then
StrWhere = StrWhere & " AND (dbo.TblAqrEarnest.StatusEarnest  =0 or dbo.TblAqrEarnest.StatusEarnest  =1 or dbo.TblAqrEarnest.StatusEarnest  =2)"
ElseIf Me.ChkOrboon(0).value = vbChecked And Me.ChkOrboon(1).value = vbChecked Then
StrWhere = StrWhere & " AND (dbo.TblAqrEarnest.StatusEarnest  =0 or dbo.TblAqrEarnest.StatusEarnest  =1)"
ElseIf Me.ChkOrboon(0).value = vbChecked And Me.ChkOrboon(2).value = vbChecked Then
StrWhere = StrWhere & " AND (dbo.TblAqrEarnest.StatusEarnest  =0 or dbo.TblAqrEarnest.StatusEarnest  =2)"
ElseIf Me.ChkOrboon(1).value = vbChecked And Me.ChkOrboon(2).value = vbChecked Then
StrWhere = StrWhere & " AND (dbo.TblAqrEarnest.StatusEarnest  =1 or dbo.TblAqrEarnest.StatusEarnest  =2)"
ElseIf Me.ChkOrboon(0).value = vbChecked Then
StrWhere = StrWhere & " AND dbo.TblAqrEarnest.StatusEarnest  =0"
ElseIf Me.ChkOrboon(1).value = vbChecked Then
StrWhere = StrWhere & " AND dbo.TblAqrEarnest.StatusEarnest  =1"
ElseIf Me.ChkOrboon(2).value = vbChecked Then
StrWhere = StrWhere & " AND dbo.TblAqrEarnest.StatusEarnest  =2"
Else
StrWhere = StrWhere & " AND (dbo.TblAqrEarnest.StatusEarnest  =0 or dbo.TblAqrEarnest.StatusEarnest  =1 or dbo.TblAqrEarnest.StatusEarnest  =2)"
End If
'If Me.ChkOrboon(0).value = vbChecked Then
'StrWhere = StrWhere & " AND dbo.TblAqrEarnest.StatusEarnest  =0"
'End If
'If Me.ChkOrboon(1).value = vbChecked Then
'StrWhere = StrWhere & " and dbo.TblAqrEarnest.StatusEarnest  =1"
'End If
'If Me.ChkOrboon(2).value = vbChecked Then
'StrWhere = StrWhere & " and dbo.TblAqrEarnest.StatusEarnest  =2"
'End If

   If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblAqrEarnest.RecordDate >=" & SQLDate(Me.Fromdate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblAqrEarnest.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If




    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL & " order by  dbo.TblAqrEarnest.ID "
  
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
If MSGType = 1 Then
         If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarOrboon.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarOrboon.rpt"
            
       End If
     Else
     
         If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarOrboon1.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarOrboon1.rpt"
            
       End If
    End If
           
        
            
    ' If Me.RdDept.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
     '       Else
      '      If Me.RdSuper.value = True Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
        '    Else
         '   If Me.RdFitter.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
          ' Else
             
            '        If Me.RdAll2.value = True Then
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
          '  Else
           '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            
     '      End If
      '      End If
       '     End If
        '     End If
         '   End If
          '  End If
        '    End If
           ' End If
          '  End If
       '      End If
           '
      '  End If



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
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        'xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim Total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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




Private Sub FromDate_Change()
If Fromdate.value <> "" Then
   FromDateH.value = ToHijriDate(Fromdate.value)
   End If
End Sub



Private Sub Fromdateh_LostFocus()

 VBA.Calendar = vbCalGreg
            Fromdate.value = ToGregorianDate(FromDateH.value)

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ToDate_Change()
If ToDate.value <> "" Then
   todateH.value = ToHijriDate(ToDate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()

 VBA.Calendar = vbCalGreg
            ToDate.value = ToGregorianDate(todateH.value)

End Sub






Private Sub TxtEmployeeID_Change()
DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.Text, True)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
