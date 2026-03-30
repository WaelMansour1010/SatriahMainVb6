VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmCheckSerial 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   0  'None
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   Icon            =   "FrmCheckSerial.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSerial 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1350
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "NO-1"
      Top             =   2580
      Width           =   4185
   End
   Begin VB.PictureBox PictureChkSerial 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   6045
      TabIndex        =   3
      Top             =   0
      Width           =   6045
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   375
      Left            =   990
      TabIndex        =   1
      Top             =   4050
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ButtonImage     =   "FrmCheckSerial.frx":038A
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
   Begin ImpulseButton.ISButton XPBtnCancel 
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   4050
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
      ButtonImage     =   "FrmCheckSerial.frx":0724
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   3120
      Width           =   5985
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "√œŒ· «·—ﬁ„ «·Œ«’ »«·„Ê“⁄ «·–Ï «Œ–  „‰Â «·‰”Œ…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   1380
      TabIndex        =   4
      Top             =   2310
      Width           =   4125
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   5580
      Picture         =   "FrmCheckSerial.frx":0ABE
      Top             =   2685
      Width           =   240
   End
End
Attribute VB_Name = "FrmCheckSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Msg As String
    CenterForm Me

    If Dir(App.path & "\Garphics\About.bmp") <> "" Then
        PictureChkSerial.Picture = LoadPicture(App.path & "\Garphics\About.bmp")
    End If

    Msg = "„·ÕÊŸ…:- ·ﬂ· „Ê“⁄ „‰ „Ê“⁄Ï «·»—‰«„Ã —ﬁ„ ”Ì—Ì«· Œ«’ »Â "
    Msg = Msg & "«œŒ· «·—ﬁ„ «·–Ï «Œœ Â „‰ «·„Ê“⁄ ... Ê≈–« ·„ Ìﬂ‰ „⁄ﬂ Ì„ﬂ‰ﬂ «‰   ’· »Â "
    Msg = Msg & "Ê≈–« ﬂ‰  ﬁ„  » Õ„Ì· «·»—‰«„Ã „‰ ⁄·Ï „Êﬁ⁄ «·‘—ﬂ… "
    Msg = Msg & "”Ê›  Ãœ «‰ Â–« «·—ﬁ„ ﬁœ «—”· ·ﬂ ⁄·Ï «·√Ì„Ì· «·Œ«’ »ﬂ "
    Me.lbl(1).Caption = Msg

End Sub

Private Sub TxtSerial_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then
        XPBtnOK_Click
    End If

End Sub

Private Sub XPBtnCancel_Click()
    Dim Msg As String
    Dim IntRes As Integer
    Msg = "≈–« ﬁ„  »€·ﬁ Â–Â «·‘«‘… «·√‰ "
    Msg = Msg & Chr(13) & "œÊ‰ «‰  œŒ· —ﬁ„ «·„Ê“⁄"
    Msg = Msg & Chr(13) & "›”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & Chr(13) & "ﬁÂ· «‰  „ «ﬂœ „‰ «·√” „—«—"
    IntRes = MsgBox(Msg, vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntRes = vbNo Then
        Exit Sub
    End If

    CloseApplication
    End
End Sub

Private Sub XPBtnOK_Click()
    Dim Msg As String
    Dim StrSerial As String
    Dim StrLeftString As String

    If Trim(TxtSerial.text) = "" Then
        Msg = "„‰ ›÷·ﬂ √œŒ· «·—ﬁ„ «·”—Ì"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
    
        StrLeftString = Trim(TxtSerial.text)

        If StrLeftString <> "NO-1" And StrLeftString <> "IB-1" And StrLeftString <> "SM-1" And StrLeftString <> "MO-1" Then
            Msg = "⁄›Ê« ...Â–« «·—ﬁ„ €Ì— ’ÕÌÕ"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        Else
            SaveSetting StrAppRegPath, "Publisher", "Publisher Key", StrLeftString
        End If
    
    End If

    Unload Me
End Sub
