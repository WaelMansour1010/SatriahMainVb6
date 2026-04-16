VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmConect_US 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·« ’«· »«·‘—ﬂ…"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "FrmConectUS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   5610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox XPTxtMobile 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1374
      Width           =   2955
   End
   Begin VB.TextBox XPTxtComName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   2955
   End
   Begin VB.TextBox XPMTxtMsg 
      Alignment       =   1  'Right Justify
      Height          =   1545
      Left            =   180
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3120
      Width           =   5205
   End
   Begin VB.TextBox XPTxtMail 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1731
      Width           =   2955
   End
   Begin VB.TextBox XPTxtPhone 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1017
      Width           =   2955
   End
   Begin VB.TextBox XPTxtVersion 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1470
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2088
      Width           =   2955
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   615
      Left            =   -30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   5655
      _cx             =   9975
      _cy             =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "«·« ’«· »«·‘—ﬂ…"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
   Begin MSComCtl2.DTPicker XPDtbSendDate 
      Height          =   285
      Left            =   2580
      TabIndex        =   5
      Top             =   2445
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Format          =   104988673
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton CmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4740
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈—”«·"
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
      ButtonImage     =   "FrmConectUS.frx":038A
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton XPBtnClear 
      Height          =   375
      Left            =   1050
      TabIndex        =   8
      Top             =   4740
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ButtonImage     =   "FrmConectUS.frx":0724
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   4740
      Width           =   795
      _ExtentX        =   1402
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
      ButtonImage     =   "FrmConectUS.frx":0ABE
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
      Caption         =   " «—ÌŒ «·«—”«·"
      Height          =   315
      Index           =   6
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2430
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·‰”Œ…"
      Height          =   315
      Index           =   5
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2076
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»—Ìœ «·«·ﬂ —Ê‰Ì"
      Height          =   315
      Index           =   4
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1722
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÃÊ«·"
      Height          =   315
      Index           =   3
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1368
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Â« ›"
      Height          =   315
      Index           =   2
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1014
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·—«”·"
      Height          =   315
      Index           =   0
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   660
      Width           =   1245
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—”«·…"
      Height          =   345
      Index           =   1
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2910
      Width           =   1065
   End
   Begin VB.Image Img 
      Height          =   1080
      Left            =   150
      Picture         =   "FrmConectUS.frx":0E58
      Top             =   780
      Width           =   1080
   End
End
Attribute VB_Name = "FrmConect_US"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim Msg As String
    Dim mail As Long
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtMail.text = "" Then
        Msg = "√œŒ· «·»—Ìœ «·«·ﬂ —Ê‰ «·Œ«’ »ﬂ" & Chr(13)
        Msg = Msg + "Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtMail.SetFocus
        Exit Sub
    End If

    mail = InStr(1, XPTxtMail.text, "@")

    If mail = 0 Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «·»—Ìœ «·«·ﬂ —Ê‰Ì " & Chr(13)
        Msg = Msg + "«·Œ«’ »ﬂ »’Ê—… ’ÕÌÕ…" & Chr(13)
        Msg = Msg + "„À· YourID@Hotmail.Com  "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtMail.SetFocus
        Exit Sub
    End If

    If XPMTxtMsg.text = "" Then
        Msg = "·„ Ì „ ﬂ «»… „Õ ÊÏ «·—”«·…" & Chr(13)
        Msg = Msg + "√œŒ· „Õ ÊÏ «·—”«·… Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPMTxtMsg.SetFocus
        Exit Sub
    End If

    Cn.BeginTrans
    BeginTrans = True
    rs.AddNew
    rs("SenderID") = CStr(new_id("SendMessage", "SenderID", "", True))
    rs("SenderName") = IIf(XPTxtComName.text = "", "", Trim(XPTxtComName.text))
    rs("Senderphone") = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
    rs("SenderMobile") = IIf(XPTxtMobile.text = "", "", Trim(XPTxtMobile.text))
    rs("SenderMail") = IIf(XPTxtMail.text = "", "", Trim(XPTxtMail.text))
    rs("VersionNum") = IIf(XPTxtVersion.text = "", "", Trim(XPTxtVersion.text))
    rs("SendDate") = IIf(XPDtbSendDate.value = "", "", XPDtbSendDate.value)
    rs("Message") = IIf(XPMTxtMsg.text = "", "", Trim(XPMTxtMsg.text))
    rs.update
    Cn.CommitTrans
    BeginTrans = False
    Unload Me
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdExit_Click
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Resize_Form Me
    XPDtbSendDate.value = Date
    Set rs = New ADODB.Recordset
    rs.Open "[SendMessage]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
        XPTxtComName.text = IIf(IsNull(rs("SenderName").value), "", Trim(rs("SenderName").value))
        XPTxtPhone.text = IIf(IsNull(rs("Senderphone").value), "", Trim(rs("Senderphone").value))
        XPTxtMobile.text = IIf(IsNull(rs("SenderMobile").value), "", Trim(rs("SenderMobile").value))
        XPTxtMail.text = IIf(IsNull(rs("SenderMail").value), "", Trim(rs("SenderMail").value))
        XPTxtVersion.text = IIf(IsNull(rs("VersionNum").value), "", Trim(rs("VersionNum").value))
        XPDtbSendDate.value = IIf(IsNull(rs("SendDate").value), "", Trim(rs("SendDate").value))
        XPMTxtMsg.text = IIf(IsNull(rs("Message").value), "", Trim(rs("Message").value))
    Else
        XPTxtComName.text = ""
        XPTxtPhone.text = ""
        XPTxtMobile.text = ""
        XPTxtMail.text = ""
        XPTxtVersion.text = ""
        XPDtbSendDate.value = Date
        XPMTxtMsg.text = ""
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Img_Click()

End Sub

Private Sub XPBtnClear_Click()
    clear_all Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

