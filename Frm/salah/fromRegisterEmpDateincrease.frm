VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FromRegisterEmpDateIncrease 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ·  «—ÌŒ  "
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
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
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   4
      Top             =   930
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
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2430
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   930
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
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      Format          =   94044161
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateExpEkamaH 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ÂÃ—Ì"
      Height          =   255
      Index           =   9
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·“Ì«œ…"
      Height          =   255
      Index           =   6
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5400
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "FromRegisterEmpDateIncrease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public fg As VSFlex8UCtl.vsFlexGrid

'Public LngRow As Long

'Public LngCol As Long

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim dateenter As Date
    Dim timEnter As Date
    Dim Askinterval As String

    If Not frmEmpSalaryComponentIncres.VSFlexGrid1 Is Nothing Then
   ' FrmCarAuthontication.fg.TextMatrix(FrmCarAuthontication.LngRow, FrmCarAuthontication.LngCol) = XPDtbBill.value    'Trim$(Me.TxtComment.text)
frmEmpSalaryComponentIncres.VSFlexGrid1.TextMatrix(frmEmpSalaryComponentIncres.LngRow, frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("EntIncresDataM")) = XPDtbBill.value
frmEmpSalaryComponentIncres.VSFlexGrid1.TextMatrix(frmEmpSalaryComponentIncres.LngRow, frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("EntIncresDataH")) = Txt_DateExpEkamaH.value
        'If lbl(8).Caption = 0 Then
        '    Askinterval = "D"
        'ElseIf lbl(8).Caption = 1 Then
        '    Askinterval = "M"
        'ElseIf lbl(8).Caption = 2 Then
           Askinterval = "dd/mm/yyyy"
        'End If
        If frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("dateenter") <> -1 Then
         '  dateenter = DateAdd(Askinterval, -2, XPDtbBill.value)
           frmEmpSalaryComponentIncres.VSFlexGrid1.TextMatrix(FrmCarAuthontication.LngRow, frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("EntIncresDataM")) = XPDtbBill.value ' dateenter
        End If
   If frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("timEnter") <> -1 Then
   '         dateenter = timeEnter ' DateAdd(Askinterval, lbl(7).Caption, timeEnter.value)
         frmEmpSalaryComponentIncres.VSFlexGrid1.TextMatrix(frmEmpSalaryComponentIncres.LngRow, frmEmpSalaryComponentIncres.VSFlexGrid1.ColIndex("EntIncresDataH")) = Txt_DateExpEkamaH.value  'timEnter
       End If
        Unload Me
    End If

End Sub
Private Sub ChangeLang()
    CmdCancel.Caption = "Cancel"
CmdOk.Caption = "Save"
lbl(6).Caption = "DateEnter"
lbl(9).Caption = "TimeEnter"
Me.Caption = "Register Date & Time"
End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
    XPDtbBill.value = Date
'Me.timeEnter.value = Time
If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub Txt_DateExpEkamaH_LostFocus()

     
      XPDtbBill.value = ToGregorianDate(Txt_DateExpEkamaH.value)
       

End Sub

Private Sub XPDtbBill_Change()
     
         Txt_DateExpEkamaH.value = ToHijriDate(XPDtbBill.value)
       

End Sub
