VERSION 5.00
Begin VB.UserControl NourHijriCal 
   AutoRedraw      =   -1  'True
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   2250
   ToolboxBitmap   =   "NourHijriCal.ctx":0000
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1860
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox TxtYear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      MaxLength       =   4
      TabIndex        =   3
      Top             =   0
      Width           =   585
   End
   Begin VB.TextBox TxtMonth 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      MaxLength       =   2
      TabIndex        =   1
      Top             =   0
      Width           =   435
   End
   Begin VB.TextBox TxtDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label LblSep 
      Alignment       =   2  'Center
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   105
   End
   Begin VB.Label LblSep 
      Alignment       =   2  'Center
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1170
      TabIndex        =   2
      Top             =   0
      Width           =   75
   End
End
Attribute VB_Name = "NourHijriCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim IntOldDay As Integer
Dim IntOldMonth As Integer
Dim IntOldYear As Integer
Dim m_ShowCheckBox As Boolean
Dim m_Checked  As Boolean
Dim m_BackColor As Single

Private Sub Chk_Click()
    TxtDay.Enabled = CBool(Chk.value)
    TxtMonth.Enabled = CBool(Chk.value)
    TxtYear.Enabled = CBool(Chk.value)
    m_Checked = CBool(Chk.value)
End Sub

Private Sub TxtDay_GotFocus()
    IntOldDay = val(TxtDay.text)
    SelectText TxtDay
End Sub

Private Sub TxtDay_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then  'Or Len(TxtDay.Text) = 2
        TxtMonth.SetFocus
    End If

End Sub

Private Sub TxtDay_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyEscape Then
        Exit Sub
    End If

    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If

End Sub

Private Sub TxtDay_LostFocus()

    If val(TxtDay.text) > 30 Then
        TxtDay.text = IntOldDay
    ElseIf Len(TxtDay.text) = 1 Then
        TxtDay.text = "0" & TxtDay.text
    ElseIf val(TxtDay.text) = 0 And IntOldDay <> 0 Then
        TxtDay.text = IntOldDay
    ElseIf Trim(TxtDay.text) = "" Then
        TxtDay.text = GetDay
    End If

End Sub

Private Sub TxtMonth_GotFocus()
    IntOldMonth = val(TxtMonth.text)
    SelectText TxtMonth
End Sub

Private Sub TxtMonth_KeyDown(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyReturn Then
        TxtYear.SetFocus
    End If

End Sub

Private Sub TxtMonth_LostFocus()

    If val(TxtMonth.text) > 12 Then
        TxtMonth.text = 12
    ElseIf Len(TxtMonth.text) = 1 Then
        TxtMonth.text = "0" & TxtMonth.text

    End If

End Sub

Private Sub TxtYear_GotFocus()
    IntOldYear = val(TxtYear.text)
    SelectText TxtYear
End Sub

Private Sub TxtYear_KeyDown(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TxtYear_LostFocus()

    If Len(Trim(TxtYear.text)) = 2 Then
        TxtYear.text = "14" & TxtYear.text
    End If

End Sub
Private Sub UserControl_Change()

End Sub
Private Sub UserControl_EnterFocus()

    If TxtDay.Enabled = True Then
        TxtDay.SetFocus
    End If

End Sub

Private Sub UserControl_Initialize()
    TxtDay.RightToLeft = True
    TxtMonth.RightToLeft = True
    TxtYear.RightToLeft = True
    'UserControl.BackColor = UserControl.Parent.BackColor
    Chk.backcolor = UserControl.backcolor
    LblSep(0).backcolor = UserControl.backcolor
    LblSep(1).backcolor = UserControl.backcolor
    UpdateLayout
    value = "Date"
End Sub

Private Sub UserControl_Resize()
    UpdateLayout
End Sub

Public Property Get ShowCheckBox() As Boolean
    ShowCheckBox = m_ShowCheckBox
End Property

Public Property Let ShowCheckBox(ByVal vNewValue As Boolean)
    m_ShowCheckBox = vNewValue

    If m_ShowCheckBox = True Then
        Chk.Visible = True
    Else
        Chk.Visible = False
    End If

    UpdateLayout
End Property

Public Property Get Checked() As Boolean

    If Me.ShowCheckBox = False Then
        'Err.Raise 513, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", "«·‘Ìﬂ »Êﬂ” €Ì— Ÿ«Â—"
    Else
        Checked = m_Checked
    End If

End Property

Public Property Let Checked(ByVal vNewValue As Boolean)

    If Me.ShowCheckBox = False Then
        Err.Raise 513, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", "«·‘Ìﬂ »Êﬂ” €Ì— Ÿ«Â—"
    Else
        m_Checked = vNewValue
        Chk.value = IIf(m_Checked = True, vbChecked, vbUnchecked)
        Chk_Click
    End If

End Property

Private Sub UpdateLayout()
    Dim SngWidth As Single
    Dim SngHeight As Single

    Dim SngAvil As Single
    Dim SngSepWidth As Single
    Dim SngChkWidth As Single

    LblSep(0).Width = 75
    LblSep(1).Width = 75
    SngSepWidth = LblSep(0).Width + LblSep(1).Width

    SngWidth = UserControl.ScaleWidth
    SngHeight = UserControl.ScaleHeight

    If Me.ShowCheckBox = False Then
        SngChkWidth = 0
    Else
        SngChkWidth = 195
    End If

    If SngWidth < (SngSepWidth) Then
        SngWidth = SngSepWidth
    End If

    SngAvil = SngWidth - (SngSepWidth + SngChkWidth)

    With TxtYear
        .left = 0
        .top = 0
        .Height = SngHeight
        .Width = 0.5 * SngAvil
    End With

    With LblSep(1)
        .left = TxtYear.left + TxtYear.Width
        .top = 0
        .Height = SngHeight
        .Width = .Width
    End With

    With TxtMonth
        .left = LblSep(1).left + LblSep(1).Width
        .top = 0
        .Height = SngHeight
        .Width = 0.25 * SngAvil
    End With

    With LblSep(0)
        .left = TxtMonth.left + TxtMonth.Width
        .top = 0
        .Height = SngHeight
        .Width = .Width
    End With

    With TxtDay
        .left = LblSep(0).left + LblSep(0).Width
        .top = 0
        .Height = SngHeight
        .Width = 0.25 * SngAvil
    End With

    If Me.ShowCheckBox = True Then

        With Chk
            .left = TxtDay.left + TxtDay.Width
            .top = 0
            .Height = SngHeight
            .Width = SngChkWidth
        End With

    End If

End Sub

Private Function GetDay() As String
    Dim StrTemp As String
    Calendar = vbCalHijri
    StrTemp = Day(Date)
    Calendar = vbCalGreg

    If Len(StrTemp) = 1 Then
        StrTemp = "0" & StrTemp
    End If

    GetDay = StrTemp
End Function

Private Sub UserControl_Show()

    If val(TxtDay.text) = 0 Then
        TxtDay.text = GetDay
    End If

    If val(TxtMonth.text) = 0 Then
        TxtMonth.text = GetMonth
    End If

    If val(TxtYear.text) = 0 Then
        TxtYear.text = GetYear
    End If

End Sub

Private Function GetMonth() As String
    Dim StrTemp As String
    Calendar = vbCalHijri
    StrTemp = Month(Date)
    Calendar = vbCalGreg

    If Len(StrTemp) = 1 Then
        StrTemp = "0" & StrTemp
    End If

    GetMonth = StrTemp
End Function

Private Function GetYear() As String
    Calendar = vbCalHijri
    GetYear = year(Date)
    Calendar = vbCalGreg
End Function

Private Sub SelectText(Txt As TextBox)

    If Len(Txt.text) > 0 Then

        With Txt
            .SelStart = 0
            .SelLength = Len(Txt.text) + 1
        End With

    End If

End Sub

Public Property Get backcolor() As Single
Attribute backcolor.VB_ProcData.VB_Invoke_Property = "StandardColor;Appearance"
Attribute backcolor.VB_UserMemId = -501
    backcolor = m_BackColor
End Property

Public Property Let backcolor(ByVal vNewValue As Single)
    m_BackColor = vNewValue
    UserControl.backcolor = vNewValue
    LblSep(0).backcolor = vNewValue
    LblSep(1).backcolor = vNewValue
    Chk.backcolor = vNewValue
End Property

Public Property Get value() As String
    Dim StrTemp As String
    StrTemp = ""
    'StrTemp = Trim(TxtDay.Text)
    'StrTemp = StrTemp & "/" & Trim(TxtMonth.Text)
    'StrTemp = StrTemp & "/" & Trim(TxtYear.Text)
    StrTemp = Trim(TxtYear.text)
    StrTemp = StrTemp & "/" & Trim(TxtMonth.text)
    StrTemp = StrTemp & "/" & Trim(TxtDay.text)
    value = StrTemp
End Property

Public Property Let value(ByVal vNewValue As String)
    On Error Resume Next
    Dim VarTemp As Variant
 
    If vNewValue = "Date" Then
        TxtDay.text = GetDay
        TxtMonth.text = GetMonth
        TxtYear.text = GetYear
    ElseIf vNewValue <> "" Then
'vNewValue = Format(vNewValue, "DD-MM-YYYY")
'        If IsDate(vNewValue) Then
'            vNewValue = Format(vNewValue, "DD-MM-YYYY")
'        End If
If InStr(1, vNewValue, "-") = 0 Then
   VarTemp = Split(vNewValue, "/", , vbTextCompare)
Else
   VarTemp = Split(vNewValue, "-", , vbTextCompare)
End If


       ' VarTemp = Split(vNewValue, "-", , vbTextCompare)

        If Len(CStr(VarTemp(0))) = 2 Then
            TxtDay.text = VarTemp(0)
        Else
            TxtYear.text = VarTemp(0)
        End If

        TxtMonth.text = VarTemp(1)
    
        If Len(CStr(VarTemp(2))) = 4 Then
            TxtYear.text = VarTemp(2)
        Else
            TxtDay.text = VarTemp(2)
        End If
    End If

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    UserControl.Enabled = vNewValue
End Property

