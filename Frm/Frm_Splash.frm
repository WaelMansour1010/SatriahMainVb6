VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5490
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   0
      Picture         =   "Frm_Splash.frx":0000
      ScaleHeight     =   3885
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   -60
         TabIndex        =   3
         Top             =   3660
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰”Œ… «·„⁄œ·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê“⁄ «·„⁄ „œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Top             =   3870
         Visible         =   0   'False
         Width           =   2865
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   630
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    '
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then

            Form_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim StrPublisherKey As String
    Dim RsOption As New ADODB.Recordset
    Dim RsRecords As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap
    CenterForm Me

    If Dir(App.path & "\Garphics\splash.bmp") <> "" Then
        Picture1.Picture = LoadPicture(App.path & "\Garphics\splash.bmp")
    End If

    RsOption.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsOption("RunCount").value = val(RsOption("RunCount").value) + 1
    RsOption.update
    StrPublisherKey = GetSetting(StrAppRegPath, "Publisher", "Publisher Key", "")
    StrPublisherKey = "NO-1"
    ' ”ÃÌ· »Ì«‰«  «·„Ê“⁄
    lbl(0).Caption = "Stars Tech."
    lbl(1).Caption = "«·—Ì«÷-«·„„·ﬂÂ «·⁄—»Ì… «·”⁄ÊœÌ…"

    Select Case StrPublisherKey

        Case "IB-1"
            lbl(1).Caption = " ›Ì Ã„ÂÊ—Ì… „’— «·⁄—»Ì…"
            lbl(2).Caption = "«·‘—ﬁ ··Õ·Ê· «· ﬂ‰Ê·ÊÃÌ… Ê ﬁ‰Ì… «·„⁄·Ê„« "
            lbl(3).Caption = " ·Ì›Ê‰:- 0124933066-0121984342 "
            SystemOptions.SysTarget = ToIbrahimShakr
            Me.lbl(4).Visible = True

        Case "MO-1"
            lbl(1).Caption = "›Ì «·„„·ﬂ… «·⁄—»Ì… «·”⁄ÊœÌ…"
            lbl(2).Caption = " „Õ›ÊŸ √Õ„œ «·”œÌ”"
            lbl(3).Caption = " —ﬁ„ «·ÃÊ«· :- 0504682366"
            SystemOptions.SysTarget = ToMahfooz
            Me.lbl(4).Visible = False

        Case "SM-1"
            lbl(1).Caption = "Ã„ÂÊ—Ì… „’— «·⁄—»Ì…"
            lbl(2).Caption = "‘—ﬂ… √”Ê«ﬁ «·„” ﬁ»· ··»—„ÃÌ« "
            lbl(3).Caption = " „Õ„Ê· :- 0127554492"
            Me.lbl(4).Visible = False

        Case "BH-1"
            '«·»Õ—Ì‰
            lbl(2).Caption = "  Script MX ·’«·Õ ‘—ﬂ… " & Chr(13) & "œÊ·… «·»Õ—Ì‰"
            lbl(3).Caption = "36646773/36360340/39822819"
            SystemOptions.SysTarget = ToBahrin
            Me.lbl(4).Visible = False

        Case Else
            lbl(1).Caption = "«·—Ì«÷-«·„„·ﬂÂ «·⁄—»Ì… «·”⁄ÊœÌ…"
            '  lbl(2).Caption = "Stars Tech."
            lbl(3).Caption = " ·Ì›Ê‰:- 014030870-014033989"
            SystemOptions.SysTarget = ToNour
            Me.lbl(4).Visible = False
    End Select

    Me.lbl(4).Visible = True
    pBar.Max = 4
    Me.Visible = False
    Me.Hide
    Exit Sub

ErrTrap:
End Sub

Private Sub Picture1_Click()
    'Unload Me
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrTrap
    Static i As Integer
    i = i + 1000

    If i >= 10000 Then
        Me.pBar.value = Me.pBar.value
        Me.Timer1.Enabled = False
        Unload Me
    Else
        Me.pBar.value = Me.pBar.value + 1
    End If

    Exit Sub
ErrTrap:
End Sub
