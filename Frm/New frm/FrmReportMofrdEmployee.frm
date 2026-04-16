VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmReportMofrdEmployee 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmReportMofrdEmployee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
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
   Begin VB.CheckBox ChkVer 
      Caption         =   "⁄—÷ —√”Ì"
      Height          =   255
      Left            =   6840
      TabIndex        =   33
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2640
      TabIndex        =   22
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   6
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   106299393
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   106299393
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   5565
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   10395
      Begin VB.Frame AttFrame 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1080
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   6915
         Begin VB.ComboBox CboYear2 
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox CmbMonth2 
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboYear 
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   1695
         End
         Begin VB.ComboBox CmbMonth 
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   315
            Index           =   13
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”‰…"
            Height          =   315
            Index           =   12
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   315
            Index           =   11
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”‰…"
            Height          =   315
            Index           =   9
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   600
            Width           =   540
         End
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   375
         Left            =   4800
         TabIndex        =   39
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„ﬁ«—‰…"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox Ch 
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " ﬁ—Ì— «Ã„«·Ì«  ‘Â—Ì"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   720
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   6795
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   16
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106299393
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106299393
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
            TabIndex        =   19
            Top             =   240
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
            TabIndex        =   18
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   6960
         TabIndex        =   13
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmReportMofrdEmployee.frx":038A
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
            TabIndex        =   14
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   12
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Top             =   6000
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbEmp 
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   1440
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbMang 
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbMofrd 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbJob 
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ChYaer 
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " ﬁ—Ì— «Ã„«·Ì«  ”‰ÊÌ"
         ForeColor       =   12582912
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbContractType 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   2520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «· ⁄«ﬁœ"
         Height          =   285
         Index           =   10
         Left            =   5490
         TabIndex        =   38
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·ÊŸÌ›… „Õœœ…"
         Height          =   285
         Index           =   6
         Left            =   5760
         TabIndex        =   32
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Top             =   4440
         Width           =   6735
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
         Height          =   1050
         Index           =   8
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   4440
         Width           =   6615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·«œ«—… „Õœœ…"
         Height          =   285
         Index           =   7
         Left            =   5640
         TabIndex        =   29
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·„›—œ „Õœœ"
         Height          =   285
         Index           =   19
         Left            =   5760
         TabIndex        =   28
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·„ÊŸ› „Õœœ"
         Height          =   285
         Index           =   5
         Left            =   5760
         TabIndex        =   25
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   6000
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
         Height          =   450
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   6240
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "·›—⁄ „Õœœ"
         Height          =   195
         Index           =   0
         Left            =   6090
         TabIndex        =   4
         Top             =   720
         Width           =   675
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   6360
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   6360
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   5640
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   106299393
      CurrentDate     =   41640
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " ‘«‘…  ﬁ«—Ì— „›—œ«  «·„ÊŸ›Ì‰    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   -45
      TabIndex        =   5
      Top             =   0
      Width           =   10500
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmReportMofrdEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public indexx As Integer
Function GetAddOrDiscount(Optional ID As Integer = 0) As Integer
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim sql As String
GetAddOrDiscount = 0
sql = "Select AddOrDiscount from mofrad where id=" & ID & ""
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetAddOrDiscount = IIf(IsNull(Rs5("AddOrDiscount").value), 0, Rs5("AddOrDiscount").value)
Else
GetAddOrDiscount = 0
End If
End Function
Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
  
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
  
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
    Exit Function
End Function
Private Sub btnClear_Click()
Cmd_Click (7)
End Sub
Sub UpdateDateEMpsalary()
Dim Yar As String
Dim Manth As String
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim i As Integer
Dim str As String
sql = "Select * from emp_salary WHERE     (RecordDate IS NULL)"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
Rs5.MoveFirst
For i = 1 To Rs5.RecordCount
Yar = IIf(IsNull(Rs5("sgn").value), "", Mid(Rs5("sgn").value, 1, 4))
Manth = IIf(IsNull(Rs5("sgn").value), "", Mid(Rs5("sgn").value, 5, 6))
str = "01/" & Manth & "/" & Yar
DTPicker1.value = CDate(str)
DTPicker1.value = MonthLastDay(DTPicker1.value)
sql = "update  emp_salary set RecordDate=" & SQLDate(DTPicker1.value, True) & " where id=" & Rs5("id").value & ""
Cn.Execute sql
Rs5.MoveNext
Next i
End If
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
  '  Label1.Visible = False
  lbl(10).Caption = "Contract"
  lblCompanyname.Caption = "AL SATTARAYH"
  Me.ChYaer.Caption = "Total Yearly Reports"
ChYaer.RightToLeft = False
Ch.Caption = "Total Monthly Reports"
Ch.RightToLeft = False
ChkVer.RightToLeft = False
ChkVer.Caption = "Vertical Viewing"
Label5.Caption = "Report of Components of Employee"
Label1(0).Caption = "Branch"
lbl(5).Caption = "Employee"
   lbl(19).Caption = "Component"
  lbl(7).Caption = "Department"
  lbl(6).Caption = "Job"
 Frame8.Caption = "Priod"
 lbl(3).Caption = "From"
 lbl(14).Caption = "To"
 lbl(8).Caption = ""
 btnClear.Caption = "Clear"
 Cmd(1).Caption = "Show"
 Cmd(2).Caption = "Exit"
End Sub

Private Sub Ch_Click()
If Me.Ch.value = vbChecked Then
ChYaer.value = vbUnchecked
Rd.value = False
AttFrame.Visible = False
Frame8.Visible = True
End If
End Sub

Private Sub ChYaer_Click()
If Me.ChYaer.value = vbChecked Then
Me.Ch.value = vbUnchecked
Rd.value = False
AttFrame.Visible = False
Frame8.Visible = True
End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       indexx = 1
 GetData
            
        Case 1
       If Ch.value = vbChecked Or Me.ChYaer.value = vbChecked Or Rd.value = True Then
       If Rd.value = True Then
       If val(CmbMonth.ListIndex) <> -1 And val(CmbMonth2.ListIndex) <> -1 And val(CboYear.ListIndex) <> -1 And val(CboYear2.ListIndex) <> -1 Then
       UpdateDateEMpsalary
       SaveTotalEmpsalary
       End If
       Else
       UpdateDateEMpsalary
       SaveTotalEmpsalary
       End If
          Else
       indexx = 0
        GetData
        End If
      If Ch.value = vbChecked Or Me.ChYaer.value = vbChecked Then
      GetDataTotal
      Else
      If val(CmbMonth.ListIndex) <> -1 And val(CmbMonth2.ListIndex) <> -1 And val(CboYear.ListIndex) <> -1 And val(CboYear2.ListIndex) <> -1 Then
      GetDataComparMofrd
      End If
      End If

  Case 7
  clear_all Me
  Fromdate.value = ""
    toDate.value = ""
ChYaer.value = vbUnchecked
Ch.value = vbUnchecked
        Case 2
        
            Unload Me
            Case 3
'print_report
    End Select
End Sub
Private Sub DcbEmp_Change()
DcbEmp_Click (0)
End Sub
Private Sub DcbEmp_Click(Area As Integer)
 If val(DcbEmp.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Public Sub YearMonth()
    Dim i As Integer
    Dim IntDefIndex As Integer
    CmbMonth.Clear
    CmbMonth2.Clear
    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
        CmbMonth2.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CmbMonth2.ListIndex = Month(Date) - 1
    CboYear.Clear
    CboYear2.Clear
    For i = 2006 To 3000
        CboYear.AddItem i
        CboYear2.AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
            IntDefIndex = CboYear2.NewIndex
        End If
    Next
    CboYear.ListIndex = IntDefIndex
    CboYear2.ListIndex = IntDefIndex

End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Dcombos.Getemp_Contract_type DcbContractType
    Dcombos.GetBranches DcbBranch
    Dcombos.GetEmployees DcbEmp
    Dcombos.GetEmpDepartments DcbMang
    Dcombos.GetEmpJobsTypes Me.DcbJob
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = " SELECT mofrad_code,mofrad_name From mofrdat order by mofrad_name"
    Else
    My_SQL = " SELECT mofrad_code,mofrad_namee From mofrdat order by mofrad_namee"
    End If
    fill_combo Me.DcbMofrd, My_SQL
    YearMonth
    Fromdate.value = ""
    toDate.value = ""
               If SystemOptions.UserInterface = EnglishInterface Then
         
        SetInterface Me
        ChangeLang
        Else
      
    End If
 
    Set cSearch = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    Resize_Form Me
End Sub
Sub SaveTotalEmpsalary()
Dim rs2 As ADODB.Recordset
Dim sql As String
Dim i As Double
Dim j As Integer

Dim Rs7 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "Delete TblTempEmpSalary where 1<>-1 "
Cn.Execute sql
Dim ColumnName As String
sql = "Select * From emp_salary where  1=1  "

'''//////////
If val(Me.DcbBranch.BoundText) <> 0 And Me.DcbBranch.Text <> "" Then
sql = sql & " AND dbo.emp_salary.BranchId = " & val(Me.DcbBranch.BoundText)
End If
If val(Me.DcbEmp.BoundText) <> 0 And Me.DcbEmp.Text <> "" Then
sql = sql & " AND dbo.emp_salary.emp_id = " & val(Me.DcbEmp.BoundText)
End If

If val(Me.DcbMang.BoundText) <> 0 And Me.DcbMang.Text <> "" Then
sql = sql & " AND dbo.emp_salary.DepartmentID = " & val(Me.DcbMang.BoundText)
End If
If val(Me.DcbEmp.BoundText) <> 0 And Me.DcbEmp.Text <> "" Then
sql = sql & " AND dbo.emp_salary.emp_id    = " & val(DcbEmp.BoundText)
End If
If Not IsNull(Me.Fromdate.value) Then
                   sql = sql & " AND dbo.emp_salary.RecordDate >=" & SQLDate(Me.Fromdate.value, True) & ""
End If

    If Not IsNull(Me.toDate.value) Then
            sql = sql & " AND  dbo.emp_salary.RecordDate <=" & SQLDate(Me.toDate.value, True) & ""
     
    End If
 ''''

 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
    Rs7.Open "TblTempEmpSalary", Cn, adOpenStatic, adLockOptimistic, adCmdTable
rs2.MoveFirst
        For i = 1 To rs2.RecordCount
            For j = 1 To 40
                ColumnName = "Comp" & j

                If Not (IsNull(rs2(ColumnName).value)) Then
                            If rs2(ColumnName) > 0 Then
                                                                Rs7.AddNew
                                                                Rs7("RecordDate").value = rs2("RecordDate").value
                                                                 Rs7("EmpID").value = rs2("emp_id").value
                                                                Rs7("BranchID").value = rs2("BranchId").value
                                                                 Rs7("MofrdID").value = j
                                                                             If GetAddOrDiscount(j) = 0 Then
                                                                             Rs7("Val").value = rs2(ColumnName).value
                                                                             Else
                                                                             Rs7("Val").value = rs2(ColumnName).value * -1
                                                                             End If
                                                                             Rs7.update 'salimmmmmmmmmm
                                            
                               End If
                   'ﬂ«‰  Â‰«
                End If
             Next j
            
            rs2.MoveNext
           Next i
      End If
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetDataTotal()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 StrSQL = "SELECT     dbo.TblTempEmpSalary.EmpID, dbo.TblTempEmpSalary.Val, dbo.TblTempEmpSalary.MofrdID, dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblEmployee.Fullcode, "
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name, dbo.TblTempEmpSalary.RecordDate, dbo.mofrad.Account_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.JobTypeID,"
 StrSQL = StrSQL & "                     dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblTempEmpSalary.BranchID, dbo.TblBranchesData.branch_name,"
 StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblEmployee.ContractID, dbo.emp_contract_type.name AS Contname, dbo.emp_contract_type.NameE AS ContnameE"
 StrSQL = StrSQL & " FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.emp_contract_type RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.emp_contract_type.id = dbo.TblEmployee.ContractID ON"
 StrSQL = StrSQL & "                     dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.mofrad RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblTempEmpSalary LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblTempEmpSalary.BranchID = dbo.TblBranchesData.branch_id ON dbo.mofrad.id = dbo.TblTempEmpSalary.MofrdID ON"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_id = dbo.TblTempEmpSalary.EmpID"
 StrSQL = StrSQL & " WHERE  (1=1)  "
    BolBegine = False
    StrWhere = ""
If val(Me.DcbMofrd.BoundText) <> 0 And Me.DcbMofrd.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblTempEmpSalary.MofrdID = " & (DcbMofrd.BoundText) & ""
End If
If val(Me.DcbJob.BoundText) <> 0 And Me.DcbJob.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.JobTypeID  = " & val(DcbJob.BoundText)
End If
If val(Me.DcbContractType.BoundText) <> 0 And Me.DcbContractType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.ContractID = " & (DcbContractType.BoundText) & ""
End If


    StrSQL = StrSQL & StrWhere


 StrSQL = StrSQL & " order by   dbo.TblTempEmpSalary.EmpID  "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        Else
        Msg = "No Data"
      End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
  
 rs.MoveFirst

 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub
Public Sub GetDataComparMofrd()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim startingDate As Date
    Dim startingDate2 As Date
    startingDate = CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))
    startingDate2 = CDate("01/" & (CmbMonth2.ListIndex + 1) & "/" & (CboYear2.ListIndex + 2006))
 StrSQL = " SELECT  c.EmpID2, x.EmpID, x.YearID, c.YearID2, x.MonthID, c.MonthID2, x.Val1, x.Val2, x.Val3, x.Val4, x.Val5, x.Val6, x.Val7, x.Val8, x.Val9, x.Val10, x.Val11, x.Val12, x.Val13, x.Val14, x.Val15, x.Val16, x.Val17, x.Val18, x.Val19, "
 StrSQL = StrSQL & "                         x.Val20, x.Val21, x.Val22, x.Val23, x.Val24, x.Val25, x.Val26, x.Val27, x.Val28, x.Val29, x.Val30, x.Val31, x.Val32, x.Val33, x.Val34, x.Val35, x.Val36, x.Val37, x.Val38, x.Val39, x.Val40, c.Val101, c.Val201, c.Val301,"
 StrSQL = StrSQL & "                         c.Val401, c.Val501, c.Val601, c.Val701, c.Val801, c.Val901, c.Val1001, c.Val1101, c.Val1201, c.Val1301, c.Val1401, c.Val1501, c.Val1601, c.Val1701, c.Val1801, c.Val1901, c.Val2001, c.Val2101, c.Val2201, c.Val2301,"
 StrSQL = StrSQL & "                         c.Val2401, c.Val2501, c.Val2601, c.Val2701, c.Val2801, c.Val2901, c.Val3001, c.Val3101, c.Val3201, c.Val3301, c.Val3401, c.Val3501, c.Val3601, c.Val3701, c.Val3801, c.Val3901, c.Val4001,"
 StrSQL = StrSQL & "                         dbo.TblEmployee.Emp_Name AS HEmp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
 StrSQL = StrSQL & "                         dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
 StrSQL = StrSQL & "                         dbo.TblEmployee.BranchId AS HBranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.NumEkama, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.dean,"
 StrSQL = StrSQL & "                         dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName,"
 StrSQL = StrSQL & "                         dbo.TblEmpDepartments.DepartmentNamee , dbo.TblEmployee.ContractID, dbo.emp_CONTRACT_TYPE.Name, dbo.emp_CONTRACT_TYPE.NameE, dbo.GetBaiscSalary(x.EmpID, " & SQLDate(startingDate, True) & ")"
 StrSQL = StrSQL & "                         AS BasicSalar, dbo.GetBaiscSalary(c.EmpID2," & SQLDate(startingDate2, True) & ") AS BasicSalar2"
 StrSQL = StrSQL & " FROM            dbo.emp_contract_type RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                         dbo.TblEmployee ON dbo.emp_contract_type.id = dbo.TblEmployee.ContractID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                         dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                         dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                         dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                             (SELECT        EmpID, YearID, MonthID, SUM(Value1) AS Val1, SUM(Value2) AS Val2, SUM(Value3) AS Val3, SUM(Value4) AS Val4, SUM(Value5) AS Val5, SUM(Value6) AS Val6, SUM(Value7) AS Val7, SUM(Value8)"
 StrSQL = StrSQL & "                                                         AS Val8, SUM(Value9) AS Val9, SUM(Value10) AS Val10, SUM(Value11) AS Val11, SUM(Value12) AS Val12, SUM(Value13) AS Val13, SUM(Value14) AS Val14, SUM(Value15) AS Val15, SUM(Value16)"
 StrSQL = StrSQL & "                                                         AS Val16, SUM(Value17) AS Val17, SUM(Value18) AS Val18, SUM(Value19) AS Val19, SUM(Value20) AS Val20, SUM(Value21) AS Val21, SUM(Value22) AS Val22, SUM(Value23) AS Val23,"
 StrSQL = StrSQL & "                                                        SUM(Value24) AS Val24, SUM(Value25) AS Val25, SUM(Value26) AS Val26, SUM(Value27) AS Val27, SUM(Value28) AS Val28, SUM(Value29) AS Val29, SUM(Value30) AS Val30, SUM(Value31)"
 StrSQL = StrSQL & "                                                         AS Val31, SUM(Value32) AS Val32, SUM(Value33) AS Val33, SUM(Value34) AS Val34, SUM(Value35) AS Val35, SUM(Value36) AS Val36, SUM(Value37) AS Val37, SUM(Value38) AS Val38,"
 StrSQL = StrSQL & "                                                         SUM(Value39) AS Val39, SUM(Value40) AS Val40"
 StrSQL = StrSQL & "                                FROM            (SELECT        EmpID, YEAR(RecordDate) AS YearID, MONTH(RecordDate) AS MonthID, CASE WHEN MofrdID = 1 THEN Val ELSE 0 END AS Value1, CASE WHEN MofrdID = 2 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value2, CASE WHEN MofrdID = 3 THEN abs(Val) ELSE 0 END AS Value3, CASE WHEN MofrdID = 4 THEN abs(Val) ELSE 0 END AS Value4,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 5 THEN abs(Val) ELSE 0 END AS Value5, CASE WHEN MofrdID = 6 THEN abs(Val) ELSE 0 END AS Value6, CASE WHEN MofrdID = 7 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value7, CASE WHEN MofrdID = 8 THEN abs(Val) ELSE 0 END AS Value8, CASE WHEN MofrdID = 9 THEN abs(Val) ELSE 0 END AS Value9,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 10 THEN abs(Val) ELSE 0 END AS Value10, CASE WHEN MofrdID = 11 THEN abs(Val) ELSE 0 END AS Value11, CASE WHEN MofrdID = 12 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value12, CASE WHEN MofrdID = 13 THEN abs(Val) ELSE 0 END AS Value13, CASE WHEN MofrdID = 14 THEN abs(Val) ELSE 0 END AS Value14,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 15 THEN abs(Val) ELSE 0 END AS Value15, CASE WHEN MofrdID = 16 THEN abs(Val) ELSE 0 END AS Value16, CASE WHEN MofrdID = 17 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value17, CASE WHEN MofrdID = 18 THEN abs(Val) ELSE 0 END AS Value18, CASE WHEN MofrdID = 19 THEN abs(Val) ELSE 0 END AS Value19,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 20 THEN abs(Val) ELSE 0 END AS Value20, CASE WHEN MofrdID = 21 THEN abs(Val) ELSE 0 END AS Value21, CASE WHEN MofrdID = 22 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value22, CASE WHEN MofrdID = 23 THEN abs(Val) ELSE 0 END AS Value23, CASE WHEN MofrdID = 24 THEN abs(Val) ELSE 0 END AS Value24,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 25 THEN abs(Val) ELSE 0 END AS Value25, CASE WHEN MofrdID = 26 THEN abs(Val) ELSE 0 END AS Value26, CASE WHEN MofrdID = 27 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value27, CASE WHEN MofrdID = 28 THEN abs(Val) ELSE 0 END AS Value28, CASE WHEN MofrdID = 29 THEN abs(Val) ELSE 0 END AS Value29,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 30 THEN abs(Val) ELSE 0 END AS Value30, CASE WHEN MofrdID = 31 THEN abs(Val) ELSE 0 END AS Value31, CASE WHEN MofrdID = 32 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value32, CASE WHEN MofrdID = 33 THEN abs(Val) ELSE 0 END AS Value33, CASE WHEN MofrdID = 34 THEN abs(Val) ELSE 0 END AS Value34,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 35 THEN abs(Val) ELSE 0 END AS Value35, CASE WHEN MofrdID = 36 THEN abs(Val) ELSE 0 END AS Value36, CASE WHEN MofrdID = 37 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value37, CASE WHEN MofrdID = 38 THEN abs(Val) ELSE 0 END AS Value38, CASE WHEN MofrdID = 39 THEN abs(Val) ELSE 0 END AS Value39,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 40 THEN abs(Val) ELSE 0 END AS Value40"
 StrSQL = StrSQL & "                                                          From dbo.TblTempEmpSalary"
 StrSQL = StrSQL & "                                                           WHERE    (MONTH(RecordDate) = " & val(CmbMonth.ListIndex + 1) & ") AND (YEAR(RecordDate) = " & val(CboYear.ListIndex + 2006) & ")) AS XTABLE"
 StrSQL = StrSQL & "                               GROUP BY EmpID,YearID, MonthID) AS x INNER JOIN"
 StrSQL = StrSQL & "                             (SELECT        EmpID2, YearID2, MonthID2, SUM(Value1) AS Val101, SUM(Value2) AS Val201, SUM(Value3) AS Val301, SUM(Value4) AS Val401, SUM(Value5) AS Val501, SUM(Value6) AS Val601, SUM(Value7)"
 StrSQL = StrSQL & "                                                         AS Val701, SUM(Value8) AS Val801, SUM(Value9) AS Val901, SUM(Value10) AS Val1001, SUM(Value11) AS Val1101, SUM(Value12) AS Val1201, SUM(Value13) AS Val1301, SUM(Value14)"
 StrSQL = StrSQL & "                                                         AS Val1401, SUM(Value15) AS Val1501, SUM(Value16) AS Val1601, SUM(Value17) AS Val1701, SUM(Value18) AS Val1801, SUM(Value19) AS Val1901, SUM(Value20) AS Val2001, SUM(Value21)"
 StrSQL = StrSQL & "                                                         AS Val2101, SUM(Value22) AS Val2201, SUM(Value23) AS Val2301, SUM(Value24) AS Val2401, SUM(Value25) AS Val2501, SUM(Value26) AS Val2601, SUM(Value27) AS Val2701, SUM(Value28)"
 StrSQL = StrSQL & "                                                         AS Val2801, SUM(Value29) AS Val2901, SUM(Value30) AS Val3001, SUM(Value31) AS Val3101, SUM(Value32) AS Val3201, SUM(Value33) AS Val3301, SUM(Value34) AS Val3401, SUM(Value35)"
 StrSQL = StrSQL & "                                                         AS Val3501, SUM(Value36) AS Val3601, SUM(Value37) AS Val3701, SUM(Value38) AS Val3801, SUM(Value39) AS Val3901, SUM(Value40) AS Val4001"
 StrSQL = StrSQL & "                                FROM            (SELECT        EmpID AS EmpID2, YEAR(RecordDate) AS YearID2, MONTH(RecordDate) AS MonthID2, CASE WHEN MofrdID = 1 THEN abs(Val) ELSE 0 END AS Value1,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 2 THEN Val ELSE 0 END AS Value2, CASE WHEN MofrdID = 3 THEN abs(Val) ELSE 0 END AS Value3, CASE WHEN MofrdID = 4 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value4, CASE WHEN MofrdID = 5 THEN Val ELSE 0 END AS Value5, CASE WHEN MofrdID = 6 THEN abs(Val) ELSE 0 END AS Value6,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 7 THEN abs(Val) ELSE 0 END AS Value7, CASE WHEN MofrdID = 8 THEN Val ELSE 0 END AS Value8, CASE WHEN MofrdID = 9 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value9, CASE WHEN MofrdID = 10 THEN abs(Val) ELSE 0 END AS Value10, CASE WHEN MofrdID = 11 THEN Val ELSE 0 END AS Value11,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 12 THEN abs(Val) ELSE 0 END AS Value12, CASE WHEN MofrdID = 13 THEN abs(Val) ELSE 0 END AS Value13,"
 StrSQL = StrSQL & "                                                                                 CASE WHEN MofrdID = 14 THEN Val ELSE 0 END AS Value14, CASE WHEN MofrdID = 15 THEN abs(Val) ELSE 0 END AS Value15, CASE WHEN MofrdID = 16 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value16, CASE WHEN MofrdID = 17 THEN Val ELSE 0 END AS Value17, CASE WHEN MofrdID = 18 THEN abs(Val) ELSE 0 END AS Value18,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 19 THEN abs(Val) ELSE 0 END AS Value19, CASE WHEN MofrdID = 20 THEN Val ELSE 0 END AS Value20, CASE WHEN MofrdID = 21 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value21, CASE WHEN MofrdID = 22 THEN abs(Val) ELSE 0 END AS Value22, CASE WHEN MofrdID = 23 THEN Val ELSE 0 END AS Value23,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 24 THEN abs(Val) ELSE 0 END AS Value24, CASE WHEN MofrdID = 25 THEN abs(Val) ELSE 0 END AS Value25,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 26 THEN Val ELSE 0 END AS Value26, CASE WHEN MofrdID = 27 THEN abs(Val) ELSE 0 END AS Value27, CASE WHEN MofrdID = 28 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value28, CASE WHEN MofrdID = 29 THEN Val ELSE 0 END AS Value29, CASE WHEN MofrdID = 30 THEN abs(Val) ELSE 0 END AS Value30,"
 StrSQL = StrSQL & "                                                                                    CASE WHEN MofrdID = 31 THEN abs(Val) ELSE 0 END AS Value31, CASE WHEN MofrdID = 32 THEN Val ELSE 0 END AS Value32, CASE WHEN MofrdID = 33 THEN abs(Val)"
 StrSQL = StrSQL & "                                                                                    ELSE 0 END AS Value33, CASE WHEN MofrdID = 34 THEN abs(Val) ELSE 0 END AS Value34, CASE WHEN MofrdID = 35 THEN Val ELSE 0 END AS Value35,"
 StrSQL = StrSQL & "  CASE WHEN MofrdID = 36 THEN abs(Val) ELSE 0 END AS Value36, CASE WHEN MofrdID = 37 THEN abs(Val) ELSE 0 END AS Value37,"
 StrSQL = StrSQL & "  CASE WHEN MofrdID = 38 THEN Val ELSE 0 END AS Value38, CASE WHEN MofrdID = 39 THEN abs(Val) ELSE 0 END AS Value39, CASE WHEN MofrdID = 40 THEN abs(Val)"
 StrSQL = StrSQL & "  ELSE 0 END AS Value40"
 StrSQL = StrSQL & "  FROM            dbo.TblTempEmpSalary AS TblTempEmpSalary_1"
 StrSQL = StrSQL & "  WHERE    (MONTH(RecordDate) = " & val(CmbMonth2.ListIndex + 1) & ") AND (YEAR(RecordDate) = " & val(CboYear2.ListIndex + 2006) & ")) AS XTABLE_1"
 StrSQL = StrSQL & "  GROUP BY EmpID2,YearID2, MonthID2) AS c ON c.EmpID2 = x.EmpID ON dbo.TblEmployee.Emp_ID = x.EmpID"
 StrSQL = StrSQL & " WHERE  (1=1)  "
    BolBegine = False
    StrWhere = ""
If val(Me.DcbEmp.BoundText) <> 0 And Me.DcbEmp.Text <> "" Then
StrWhere = StrWhere & " AND x.EmpID = " & val(DcbEmp.BoundText) & ""
End If
If val(Me.DcbJob.BoundText) <> 0 And Me.DcbJob.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.JobTypeID  = " & val(DcbJob.BoundText)
End If
If val(Me.DcbContractType.BoundText) <> 0 And Me.DcbContractType.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.ContractID = " & val(DcbContractType.BoundText) & ""
End If
If val(Me.DcbBranch.BoundText) <> 0 And Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.BranchID = " & val(DcbBranch.BoundText) & ""
End If
If val(Me.DcbBranch.BoundText) <> 0 And Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.BranchID = " & val(DcbBranch.BoundText) & ""
End If

    StrSQL = StrSQL & StrWhere


 'StrSQL = StrSQL & " order by   x.EmpID  "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        Else
        Msg = "No Data"
      End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
  
 rs.MoveFirst

 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

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

 StrSQL = "SELECT     dbo.TblEmpMofrd.ID, dbo.TblEmpMofrd.RecorDate, dbo.TblEmpMofrd.EmpID, dbo.TblEmpMofrd.MofrdCode, dbo.TblEmpMofrd.Valu, dbo.TblEmpMofrd.BrnchID, "
 StrSQL = StrSQL & "                     dbo.TblEmpMofrd.MofrdName, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmpMofrd.MorfID, dbo.TblBranchesData.branch_name,"
 StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.DepartmentID , dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.NationalityE"
 StrSQL = StrSQL & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblEmpDepartments.DeparmentID = dbo.TblEmployee.DepartmentID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmpMofrd LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblEmpMofrd.BrnchID = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.TblEmpMofrd.EmpID"
 StrSQL = StrSQL & " WHERE  (1=1)  "
    BolBegine = False
    StrWhere = ""
If val(Me.DcbBranch.BoundText) <> 0 And Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)
End If
If val(Me.DcbMang.BoundText) <> 0 And Me.DcbMang.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.DepartmentID = " & val(Me.DcbMang.BoundText)
End If
If val(Me.DcbEmp.BoundText) <> 0 And Me.DcbEmp.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmpMofrd.EmpID    = " & val(DcbEmp.BoundText)
End If
If val(Me.DcbMofrd.BoundText) <> 0 And Me.DcbMofrd.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmpMofrd.MofrdName = '" & (DcbMofrd.Text) & "'"
End If
If val(Me.DcbJob.BoundText) <> 0 And Me.DcbJob.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblEmployee.JobTypeID  = " & val(DcbJob.BoundText)
End If
If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpMofrd.RecorDate >=" & SQLDate(Me.Fromdate.value, True) & ""
End If

    If Not IsNull(Me.toDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblEmpMofrd.RecorDate <=" & SQLDate(Me.toDate.value, True) & ""
     
    End If




    '-----------------------------------

    StrSQL = StrSQL & StrWhere


 StrSQL = StrSQL & " order by   dbo.TblEmpMofrd.EmpID  "

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
  
 rs.MoveFirst

 print_report StrSQL

'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Dim str2 As String
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If Ch.value = vbChecked Then
     If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmpSalary.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmpSalaryE.rpt"
            
       End If
  ElseIf Me.ChYaer.value = vbChecked Then
       If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmpSalaryYear.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmpSalaryYaerE.rpt"
            
       End If
Else

If ChkVer.value = vbChecked Then


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmployee.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmployeeE.rpt"
            
       End If
  ElseIf Rd.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdCompEmployee.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdCompEmployee.rpt"
            
       End If

Else


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmployee1.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportMofrdEmployeeE1.rpt"
            
       End If
End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
    Dim MSGType As Integer
   
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
   
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
     
    End If


   ' xReport.ParameterFields(3).AddCurrentValue user_name
   If Fromdate.value <> "" And toDate.value <> "" Then
    xReport.ParameterFields(14).AddCurrentValue Fromdate.value
       
       xReport.ParameterFields(16).AddCurrentValue toDate.value
      
       End If
       If Rd.value = True Then
       xReport.ParameterFields(19).AddCurrentValue CboYear.Text & "/" & CmbMonth.ListIndex + 1
       xReport.ParameterFields(20).AddCurrentValue CboYear2.Text & "/" & CmbMonth2.ListIndex + 1
       End If
       Dim str As String
       If Me.ChYaer.value = vbChecked Or Ch.value = vbChecked Then
       str = ""
       If SystemOptions.UserInterface = ArabicInterface Then
       str2 = " ﬁ—Ì— «·„›—œ«    "
       If val(DcbBranch.BoundText) <> 0 And DcbBranch.Text <> "" Then
       str = str & " " & "›—⁄" & " " & Me.DcbBranch.Text & " " & CHR(13)
       End If
        If val(DcbMang.BoundText) <> 0 And DcbMang.Text <> "" Then
       str = str & "  " & "«œ«—…" & " " & Me.DcbMang.Text & " " & CHR(13)
       End If
           If val(DcbEmp.BoundText) <> 0 And DcbEmp.Text <> "" Then
       str = str & "  " & "«·„ÊŸ›" & " " & Me.DcbEmp.Text & " " & CHR(13)
       End If
            If val(DcbMofrd.BoundText) <> 0 And DcbMofrd.Text <> "" Then
       str = str & "  " & "«·„›—œ" & " " & Me.DcbMofrd.Text & " " & CHR(13)
       End If
              If val(DcbJob.BoundText) <> 0 And DcbJob.Text <> "" Then
       str = str & "  " & "ÊŸÌ›…" & " " & Me.DcbJob.Text & " " & CHR(13)
       End If
       Else
         str2 = "Report Component "
       If val(DcbBranch.BoundText) <> 0 And DcbBranch.Text <> "" Then
       str = str & " " & "Branch" & " " & Me.DcbBranch.Text & " " & CHR(13)
       End If
        If val(DcbMang.BoundText) <> 0 And DcbMang.Text <> "" Then
       str = str & "  " & "Management" & " " & Me.DcbMang.Text & " " & CHR(13)
       End If
           If val(DcbEmp.BoundText) <> 0 And DcbEmp.Text <> "" Then
       str = str & "  " & "Employee" & " " & Me.DcbEmp.Text & " " & CHR(13)
       End If
            If val(DcbMofrd.BoundText) <> 0 And DcbMofrd.Text <> "" Then
       str = str & "  " & "Component" & " " & Me.DcbMofrd.Text & " " & CHR(13)
       End If
              If val(DcbJob.BoundText) <> 0 And DcbJob.Text <> "" Then
       str = str & "  " & "Job" & " " & Me.DcbJob.Text & " " & CHR(13)
       End If
       End If
        If str <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
    str2 = str2 & " »Õ”» " & str
    Else
    str2 = str2 & "By" & str
    End If
    End If
       xReport.ParameterFields(12).AddCurrentValue str2
       End If
       
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



Private Sub Rd_Click()
If Rd.value = True Then
Me.Ch.value = vbUnchecked
Me.ChYaer.value = vbUnchecked
AttFrame.Visible = True
Frame8.Visible = False
Else
Frame8.Visible = True
AttFrame.Visible = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcbEmp.BoundText = EmpID
    End If
End Sub
