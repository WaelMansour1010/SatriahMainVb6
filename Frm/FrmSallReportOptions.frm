VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSallReportOptions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ŒÌ«—«  «·ÿ»«⁄…"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "FrmSallReportOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·›« Ê—…"
      Height          =   945
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   5145
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ ”Ì—Ì«· „Œ ’—"
         Height          =   375
         Index           =   4
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1485
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ ‘Ã—Ï ··√’‰«› Ê Ã„Ì⁄ »«·√”›·"
         Height          =   255
         Index           =   3
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   930
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Frame FrmPrinter 
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÿ»«⁄…"
         Height          =   675
         Left            =   420
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1980
         Visible         =   0   'False
         Width           =   4485
         Begin VB.OptionButton OptPrinterType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ«»⁄… ‰ﬁÿÌ…"
            Height          =   375
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton OptPrinterType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ«»⁄… ﬂ«‘Ì—"
            Height          =   375
            Index           =   1
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton OptPrinterType 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»«⁄… ⁄«œÌ…"
            Height          =   375
            Index           =   0
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.TextBox TxtDataPath 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   570
         TabIndex        =   8
         Top             =   1650
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Œ’’"
         Height          =   255
         Index           =   2
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ „Œ ’—"
         Height          =   255
         Index           =   1
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   345
         Width           =   1485
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷  ›’Ì·Ì"
         Height          =   255
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   1485
      End
      Begin ImpulseButton.ISButton CmdPath 
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   609
         Caption         =   "..."
         BackColor       =   14737632
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSallReportOptions.frx":038A
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label LBL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Õœœ „”«— «· ﬁ—Ì—"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   4740
         Picture         =   "FrmSallReportOptions.frx":0724
         Top             =   1320
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.CheckBox ChkAsk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«  ”√· „—… √Œ—Ï"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1380
      Width           =   1575
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   1110
      TabIndex        =   4
      Top             =   2190
      Width           =   1035
      _ExtentX        =   1826
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
      ButtonImage     =   "FrmSallReportOptions.frx":0AAE
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
      Left            =   90
      TabIndex        =   5
      Top             =   2190
      Width           =   915
      _ExtentX        =   1614
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
      ButtonImage     =   "FrmSallReportOptions.frx":0E48
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
   Begin VB.Label LBL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "·≈ŸÂ«—Â« „—… «Œ—Ï .„‰  ‘«‘… «·ŒÌ«—« - «Œ — ŒÌ«—«  ⁄—÷-≈ŸÂ«— ŒÌ«—«  «·ÿ»«⁄…"
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
      Height          =   405
      Index           =   1
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1590
      Width           =   4875
   End
   Begin VB.Label LBL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õœœ ‰Ê⁄ «·›Ê« Ì— «·–Ì  —€» ›Ì ÿ»«⁄ Â"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1050
      Width           =   2745
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5220
      X2              =   -30
      Y1              =   2070
      Y2              =   2055
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   4920
      Picture         =   "FrmSallReportOptions.frx":11E2
      Top             =   1050
      Width           =   240
   End
End
Attribute VB_Name = "FrmSallReportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_UserCanceled As Boolean

Private Sub CmdPath_Click()
    On Error GoTo ErrTrap

    With mdifrmmain.Cmdlg
        .CancelError = False
        .DialogTitle = "Õœœ „”«— «· ﬁ—Ì—"
        .Filter = "Report Designer|*.drp"
        .InitDir = App.path & "\Bill_Template"
        .ShowOpen

        If .FileName = "" Then
        Else
            TxtDataPath.text = .FileName
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrPath  As String
    Dim PrintType As Integer

    CenterForm Me

    FormPostion Me, GetPostion
    StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")
    TxtDataPath.text = StrPath
    PrintType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

    Select Case PrintType

        Case 1
            Opt(0).value = True
            Me.OptPrinterType(0).value = True

        Case 2
            Opt(1).value = True
            Me.OptPrinterType(0).value = True

        Case 3
            Opt(2).value = True
            Me.OptPrinterType(0).value = True

        Case 4
            Opt(1).value = True
            Me.OptPrinterType(1).value = True

        Case 5
            Opt(1).value = True
            Me.OptPrinterType(0).value = True

        Case 6
            Opt(1).value = True
            Me.OptPrinterType(2).value = True
    End Select

   If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
 
    Me.Caption = "Print Options"
Frame1.Caption = "Show Type"
Opt(0).Caption = "Detailed"
Opt(1).Caption = "Shortly"
Opt(4).Caption = "Serials in one Row"
LBL(1).Caption = "To Show this Screen again - goto system options -> view Options"
ChkAsk.Caption = "Dont Ask Again"

 LBL(2).Caption = "Select Print Type"

    XPBtnOK.Caption = "Ok"
    XPBtnCancel.Caption = "Cancel"
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub Opt_Click(Index As Integer)
    CmdPath.Enabled = Opt(2).value
    LBL(0).Enabled = Opt(2).value
End Sub

Private Sub OptPrinterType_Click(Index As Integer)

    Select Case Index

        Case 0
            Opt(0).Enabled = True
            Opt(1).Enabled = True
            Opt(2).Enabled = True

            If Opt(2).value = True Then
                Image1(1).Enabled = True
                LBL(0).Enabled = True
                TxtDataPath.Enabled = True
                CmdPath.Enabled = True
            Else
                Image1(1).Enabled = False
                LBL(0).Enabled = False
                TxtDataPath.Enabled = False
                CmdPath.Enabled = False
            End If

        Case 1
            Opt(0).Enabled = False
            Opt(1).Enabled = True
            Opt(2).Enabled = False
            Opt(1).value = True
            Image1(1).Enabled = False
            LBL(0).Enabled = False
            TxtDataPath.Enabled = False
            CmdPath.Enabled = False
    End Select

End Sub

Private Sub XPBtnCancel_Click()
    On Error GoTo ErrTrap
    Me.UserCanceled = True
    Me.Hide
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    '≈ŸÂ«— √Ê ⁄œ„ ≈ŸÂ«— ‰«›–… «·ŒÌ«—« 
    Dim Msg As String

    If Opt(2).value = True And TxtDataPath.text = "" Then
        Msg = "⁄‰œ «Œ Ì«— ⁄—÷ „Œ’’" & Chr(13)
        Msg = Msg + "ÌÃ»  ÕœÌœ «· ﬁ—Ì— «·–Ì  —€» ›Ì «” Œœ«„Â"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CmdPath_Click
        Exit Sub

        If Dir(TxtDataPath.text) = "" Then
            Msg = " «· ﬁ—Ì— «·„Õœœ €Ì— „ÊÃÊœ" & Chr(13)
            Msg = Msg + " √ﬂœ „‰ «·„”«— «·’ÕÌÕ "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CmdPath_Click
            Exit Sub
        End If
    End If

    SaveSetting StrAppRegPath, "View_Type", "ShowMe", ChkAsk.value

    If TxtDataPath.text <> "" Then
        SaveSetting StrAppRegPath, "PrintReport", "ReportPath", TxtDataPath.text
    End If

    '‰Ê⁄ «·ÿ»«⁄…
    If OptPrinterType(0).value = True Then
        If Opt(0).value = True Then
            SaveSetting StrAppRegPath, "View_Type", "SallReportType", 1
        ElseIf Opt(1).value = True Then
            SaveSetting StrAppRegPath, "View_Type", "SallReportType", 2
        ElseIf Opt(2).value = True Then
            SaveSetting StrAppRegPath, "View_Type", "SallReportType", 3
        ElseIf Opt(3).value = True Then
            SaveSetting StrAppRegPath, "View_Type", "SallReportType", 5
        ElseIf Opt(4).value = True Then
            SaveSetting StrAppRegPath, "View_Type", "SallReportType", 40
        End If

    ElseIf OptPrinterType(1).value = True Then
        SaveSetting StrAppRegPath, "View_Type", "SallReportType", 4
    ElseIf OptPrinterType(2).value = True Then
        'ÿ«»⁄… ‰ﬁÿÌ…
        SaveSetting StrAppRegPath, "View_Type", "SallReportType", 6
    End If

    frmsalebill.BolPrint = True
    Unload Me
    Exit Sub
ErrTrap:
End Sub

Public Property Get UserCanceled() As Boolean
    UserCanceled = m_UserCanceled
End Property

Public Property Let UserCanceled(ByVal vNewValue As Boolean)
    m_UserCanceled = vNewValue
End Property
