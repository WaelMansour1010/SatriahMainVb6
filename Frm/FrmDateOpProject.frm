VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDateOpProject 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· «· «—ÌŒ "
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   4
      Top             =   1170
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4830
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   1170
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      Format          =   142540801
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DcTime 
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   142540802
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin MSComCtl2.DTPicker ToTime 
      Height          =   330
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   142540802
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   142540802
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”«⁄…"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ï «·”«⁄…"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ «·”«⁄…"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   1080
      Y2              =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   4380
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   1
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      TabIndex        =   0
      Top             =   1020
      Width           =   1545
   End
End
Attribute VB_Name = "FrmDateOpProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public fg As VSFlex8UCtl.vsFlexGrid

'Public LngRow As Long

Public Index As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim Period As Double
    Dim Msg As String
    Dim dateenter As Date
    Dim timEnter As Date
    Dim Askinterval As String
    Period = 0
On Error Resume Next

If Index = 0 Then

    If Not Projects.VSFlexGrid2 Is Nothing Then
If Projects.LngCol = 21 Then
 Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("StartDate")) = XPDtbBill.value
 ElseIf Projects.LngCol = 22 Then
 Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("EndDate")) = XPDtbBill.value
 End If
   Unload Me
    End If
          ElseIf Index = 540 Then
        FrmAccEditJournal.Fg_Journal.TextMatrix(FrmAccEditJournal.Fg_Journal.LngRow, FrmAccEditJournal.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me

      ElseIf Index = 541 Then
        FrmAccEditJournal1.Fg_Journal.TextMatrix(FrmAccEditJournal1.LngRow, FrmAccEditJournal1.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me
      ElseIf Index = 542 Then
        FrmAccEditJournal2.Fg_Journal.TextMatrix(FrmAccEditJournal2.LngRow, FrmAccEditJournal2.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me

      ElseIf Index = 543 Then
        FrmAccEditJournal3.Fg_Journal.TextMatrix(FrmAccEditJournal3.LngRow, FrmAccEditJournal3.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me

      ElseIf Index = 544 Then
        FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me
  
    ElseIf Index = 1 Then


    
 
   Unload Me

    End If
    
'      ElseIf Index = 2 Then
'    FrmEmpIncreaseSalaries.VSFlexGrid1.TextMatrix(FrmEmpIncreaseSalaries.LngRow, FrmEmpIncreaseSalaries.VSFlexGrid1.ColIndex("RecoedDate")) = XPDtbBill.value
'    Unload Me
'

'
'
'
        


End Sub
Private Sub ChangeLang()
    cmdCancel.Caption = "Cancel"
CMDOK.Caption = "Save"
lbl(6).Caption = "DateEnter"
' 'bl(9).Caption = "TimeEnter"
Me.Caption = "Register Date "


End Sub

Private Sub DatePicker1_SelectionChanged()
'XPDtbBill.value = DatePicker1.AttachToCalendar
End Sub

Private Sub Form_Load()
    CenterForm Me
XPDtbBill.value = Date
DcTime = Time
ToTime.value = Time
lbl(0).Visible = False
lbl(1).Visible = False
ToTime.Visible = False
DTPicker1.Visible = False
lbl(2).Visible = False
DcTime.value = ""
Me.ToTime.value = ""
XPDtbBill.value = Date
    FormPostion Me, GetPostion
XPDtbBill.Visible = True
Txt_DateHigri.Visible = True


If Index = 30 Then
DTPicker1.Visible = False
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False

 If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")) <> "" Then
   XPDtbBill.value = FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate"))
   End If

End If
If Index = 31 Then
DTPicker1.Visible = False
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False

 If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate")) <> "" Then
   XPDtbBill.value = FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate"))
   End If

End If

If Index = 29 Then
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False
End If
On Error Resume Next



   
 

If Index = 33 Or Index = 34 Then
    lbl(0).Visible = True
    lbl(0).Caption = "«·Êﬁ "
    DcTime.Visible = True
    DcTime.CheckBox = True
    
End If
If Index = 35 Then
    lbl(0).Visible = False
    
    DcTime.Visible = False
    DcTime.CheckBox = False
    
End If

If Index = 36 Or Index = 37 Or Index = 38 Or Index = 39 Then
    lbl(0).Visible = False
    
    DcTime.Visible = False
    DcTime.CheckBox = False
    
End If



    Me.CMDOK.ButtonStyle = impActive
    Set CMDOK.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CMDOK.ButtonPositionImage = impRightOfText

    Me.cmdCancel.ButtonStyle = impActive
    Set cmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    cmdCancel.ButtonPositionImage = impRightOfText
    
'Me.timeEnter.value = Time
If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub



Private Sub Txt_DateHigri_LostFocus()
VBA.Calendar = vbCalGreg
            XPDtbBill.value = ToGregorianDate(Txt_DateHigri.value)
End Sub

Private Sub XPDtbBill_Change()
If IsDate(XPDtbBill.value) Then
    Txt_DateHigri.value = ToHijriDate(XPDtbBill.value)
End If
End Sub
