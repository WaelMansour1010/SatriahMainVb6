VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmDelUser 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Õ–› „” Œœ„"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   HelpContextID   =   260
   Icon            =   "FrmDelUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1050
      TabIndex        =   1
      Top             =   1170
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
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
      ColorHighlight  =   4194304
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   1170
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   4194304
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   2490
      TabIndex        =   3
      Top             =   1170
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      ButtonImage     =   "FrmDelUser.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   1
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6330
      X2              =   -30
      Y1              =   1050
      Y2              =   1050
   End
End
Attribute VB_Name = "FrmDelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo  As clsDCboSearch

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Unload Me
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    CenterForm Me

    FormPostion Me, GetPostion
    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboUserName
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset

    If DCboUserName.BoundText = 1 Then
        Msg = "Â–« «·„” Œœ„ ÂÊ „œÌ— «·‘—ﬂ… √Ê «·„‰‘√…" & Chr(13)
        Msg = Msg + "·« Ì„ﬂ‰ Õ–› »Ì«‰«  «·„œÌ—" & Chr(13)
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    StrSQL = "select * From Transactions where UserID=" & DCboUserName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› Â–« «·„” Œœ„ " & Chr(13)
        Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  «· Ì ﬁ«„ »Â« Â–« «·„” Œœ„"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsTemp.Close
        Exit Sub
    End If

    StrSQL = "delete  From TblUsers where UserID=" & DCboUserName.BoundText
    Cn.Execute StrSQL
    Msg = " „ Õ–› Â–« «·„” Œœ„ »‰Ã«Õ"
    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–« «·„” Œœ„ " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
