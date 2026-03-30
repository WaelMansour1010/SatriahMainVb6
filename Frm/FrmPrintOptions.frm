VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPrintOptions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3075
   Icon            =   "FrmPrintOptions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkAsk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«  ”√· „—… √Œ—Ï"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1050
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1110
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·⁄—÷"
      Height          =   915
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   2595
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷  ›’Ì·Ì"
         Height          =   255
         Index           =   0
         Left            =   780
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ „Œ ’—"
         Height          =   255
         Index           =   1
         Left            =   780
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1485
      End
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   1590
      TabIndex        =   4
      Top             =   2160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmPrintOptions.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   630
      TabIndex        =   5
      Top             =   2160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmPrintOptions.frx":0724
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
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   2820
      Picture         =   "FrmPrintOptions.frx":0ABE
      Top             =   1500
      Width           =   240
   End
   Begin VB.Line Line1 
      X1              =   3060
      X2              =   30
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õœœ ‰Ê⁄ «·›Ê« Ì— «·–Ì  —€» ›Ì ÿ»«⁄ Â"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1500
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "FrmPrintOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ChangeLang()
 
    Me.Caption = "Print Options"
Frame1.Caption = "Show Type"
Opt(0).Caption = "Detailed"
Opt(1).Caption = "Shortly"
ChkAsk.Caption = "Dont Ask Again"

 Label1.Caption = "Select Print Type"

    XPBtnOK.Caption = "Ok"
    XPBtnCancel.Caption = "Cancel"
End Sub

Private Sub Command1_Click()
    MsgBox mdifrmmain.ActiveForm.name
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        Opt(0).value = True
    Else
        Opt(1).value = True
    End If

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnCancel_Click()
    On Error GoTo ErrTrap

    If mdifrmmain.ActiveForm Is Nothing Then Unload Me
    mdifrmmain.ActiveForm.BolPrint = False
    Unload Me
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnOK_Click()
    SaveSetting StrAppRegPath, "View_Type", "ShowMe", ChkAsk.value
    SaveSetting StrAppRegPath, "View_Type", "ReportType", Opt(0).value
    Unload Me
End Sub
