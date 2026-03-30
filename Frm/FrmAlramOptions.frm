VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAlramOptions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ŒÌ«—«   ‰»ÌÂ «·√Ê—«ﬁ «·„«·Ì… «·„” Õﬁ…"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "FrmAlramOptions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ‘€Ì· ’Ê  «· ‰ÌÂ ⁄‰œ ŸÂÊ—Â"
      Height          =   255
      Index           =   1
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1740
      Width           =   2595
   End
   Begin VB.TextBox TxtPath 
      Height          =   345
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   3915
   End
   Begin ImpulseButton.ISButton CmdColor 
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   7
      Top             =   510
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmAlramOptions.frx":038A
      DrawFocusRectangle=   0   'False
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ›⁄Ì· «·√·Ê«‰ "
      Height          =   345
      Index           =   0
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2595
   End
   Begin ImpulseButton.ISButton CmdColor 
      Height          =   345
      Index           =   1
      Left            =   540
      TabIndex        =   8
      Top             =   870
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmAlramOptions.frx":0724
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdColor 
      Height          =   345
      Index           =   2
      Left            =   540
      TabIndex        =   9
      Top             =   1230
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmAlramOptions.frx":0ABE
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdColor 
      Height          =   345
      Index           =   3
      Left            =   30
      TabIndex        =   12
      Top             =   2400
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmAlramOptions.frx":0E58
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3420
      Width           =   825
      _ExtentX        =   1455
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
      ButtonImage     =   "FrmAlramOptions.frx":11F2
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   375
      Left            =   1020
      TabIndex        =   17
      Top             =   3420
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      ButtonImage     =   "FrmAlramOptions.frx":158C
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   5
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2790
      Width           =   3675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·ÕÊŸ…:-"
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
      Height          =   255
      Index           =   4
      Left            =   3750
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”«— „·› «·’Ê  «·Œ«’ »«· ‰»ÌÂ"
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
      Height          =   255
      Index           =   3
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2100
      Width           =   2325
   End
   Begin VB.Label LblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label LblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   510
      Width           =   840
   End
   Begin VB.Label LblColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   900
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4560
      X2              =   30
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê—ﬁ… „«·Ì… ﬁ—» „Ì⁄«œ ≈” Õﬁ«ﬁÂ«"
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
      Height          =   315
      Index           =   2
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1230
      Width           =   2415
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê—ﬁ… „«·Ì…  «—ÌŒ ≈” Õﬁ«ﬁÂ« «·ÌÊ„"
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
      Height          =   315
      Index           =   1
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   2415
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ê—ﬁ… „«·Ì… ›«   «—ÌŒ ≈” Õﬁ«ﬁÂ«"
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
      Height          =   315
      Index           =   0
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   2415
   End
End
Attribute VB_Name = "FrmAlramOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_Click(Index As Integer)

    Select Case Index

        Case 0
            Me.lbl(0).Enabled = CBool(Me.Chk(Index).value)
            Me.lbl(1).Enabled = CBool(Me.Chk(Index).value)
            Me.lbl(2).Enabled = CBool(Me.Chk(Index).value)
        
            Me.LblColor(0).Enabled = CBool(Me.Chk(Index).value)
            Me.LblColor(1).Enabled = CBool(Me.Chk(Index).value)
            Me.LblColor(2).Enabled = CBool(Me.Chk(Index).value)
        
            Me.CmdColor(0).Enabled = CBool(Me.Chk(Index).value)
            Me.CmdColor(1).Enabled = CBool(Me.Chk(Index).value)
            Me.CmdColor(2).Enabled = CBool(Me.Chk(Index).value)

        Case 1
            Me.lbl(3).Enabled = CBool(Me.Chk(Index).value)
            Me.TxtPath.Enabled = CBool(Me.Chk(Index).value)
            Me.CmdColor(3).Enabled = CBool(Me.Chk(Index).value)
        
    End Select

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdColor_Click(Index As Integer)

    With mdifrmmain.cmDlg
        .CancelError = False

        Select Case Index

            Case 0
                .color = Me.LblColor(0).BackColor
                .ShowColor
                Me.LblColor(0).BackColor = .color

            Case 1
                .color = Me.LblColor(1).BackColor
                .ShowColor
                Me.LblColor(1).BackColor = .color

            Case 2
                .color = Me.LblColor(2).BackColor
                .ShowColor
                Me.LblColor(2).BackColor = .color

            Case 3
                .Filter = "Wave Files(*.wav)|*.wav"
                .DialogTitle = "≈Œ Ì«— „·› ’Ê  «· ‰»ÌÂ"
                .ShowOpen

                If .FileName <> "" Then
                    Me.TxtPath.text = .FileName
                End If

        End Select

    End With

End Sub

Private Sub CmdOk_Click()
    Dim Msg As String

    If Me.TxtPath.text <> "" Then
        If Dir(Me.TxtPath.text, vbNormal) = "" Then
            Msg = "„”«— „·› «·’Ê  «·–Ï «œŒ· Â €Ì— ’ÕÌÕ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If

    If Get_Setting(1) = True Then
        Msg = " „ Õ›Ÿ Â–Â «·√⁄œ«œ  "
        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        FrmPaymentTime.ApplySetting
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    Dim Msg As String
    Msg = "›Ï Õ«·… ⁄„· «·»—‰«„Ã ⁄·Ï «·‘»ﬂ… ›«‰Â ÌÃ»  ÕœÌœ Â–« «·ŒÌ«—"
    Msg = Msg & "⁄·Ï ﬂ· ÃÂ«“ „‰ «·„ÊÃÊœÌ‰ ⁄·Ï «·‘»ﬂ…"
    Me.lbl(5).Caption = Msg
    CenterForm Me

    FormPostion Me, GetPostion
    Get_Setting 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Function Get_Setting(IntMode As Integer) As Boolean
    'IntMode=0 Read Setting
    'IntMode=1 Write Setting

    Dim rs As ADODB.Recordset

    On erorr GoTo ErrTrap

    Set rs = New ADODB.Recordset

    If IntMode = 0 Then
        rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

        If rs("EnableNotesAlramColors").value = 1 Then
            Me.Chk(0).value = vbChecked
        Else
            Me.Chk(0).value = vbUnchecked
        End If

        Me.LblColor(0).BackColor = IIf(IsNull(rs("Color1").value), vbWhite, rs("Color1").value)
        Me.LblColor(1).BackColor = IIf(IsNull(rs("Color2").value), vbWhite, rs("Color2").value)
        Me.LblColor(2).BackColor = IIf(IsNull(rs("Color3").value), vbWhite, rs("Color3").value)
        Chk_Click 0

        If rs("PlayNotesAlramSound").value = 1 Then
            Me.Chk(1).value = vbChecked
        Else
            Me.Chk(1).value = vbUnchecked
        End If

        Me.TxtPath.text = IIf(IsNull(rs("AlramSoundFilePath").value), "", rs("AlramSoundFilePath").value)
        Chk_Click 1
    ElseIf IntMode = 1 Then
        rs.Open "TblOptions", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        rs("EnableNotesAlramColors").value = Me.Chk(0).value
        rs("Color1").value = Me.LblColor(0).BackColor
        rs("Color2").value = Me.LblColor(1).BackColor
        rs("Color3").value = Me.LblColor(2).BackColor
    
        rs("PlayNotesAlramSound").value = Me.Chk(1).value
        rs("AlramSoundFilePath").value = Me.TxtPath.text
        rs.update
        Get_Setting = True
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    Get_Setting = False
End Function

