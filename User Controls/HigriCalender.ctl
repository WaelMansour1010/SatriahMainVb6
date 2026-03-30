VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl HigriCalender 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2580
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   2580
   ToolboxBitmap   =   "HigriCalender.ctx":0000
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   390
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2580
      _cx             =   4551
      _cy             =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   800
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   2
      ChildSpacing    =   2
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
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
      Begin VB.TextBox TxtYear 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   45
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   45
         Width           =   570
      End
      Begin VB.TextBox TxtDay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   2100
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   45
         Width           =   420
      End
      Begin VB.ComboBox CboMonth 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   630
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   45
         Width           =   1455
      End
   End
End
Attribute VB_Name = "HigriCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Change()
Dim m_Value As String
Dim Old_Day As String
Dim Old_Year As String
Dim m_BorderColor As Long
Public Enum HijriMonths
    Moharam ' „Õ—„
    Safar '’›—
    RabFirs '—»Ì⁄ √Ê·
    RabSecond '—»Ì⁄ À«‰Ï
    GamadaFirst 'Ã„«œÏ √Ê·
    GamadaSecond 'Ã„«œÏ À«‰Ï
    Ragab '—Ã»
    Shaaban '‘⁄»«‰
    Ramadan '—„÷«‰
    Shawal '‘Ê«·
    ZoAlKada '–Ê «·ﬁ⁄œ…
    ZoAlHaga '–Ê «·ÕÃ…
End Enum
Private Sub CboMonth_Change()
RaiseEvent Change
End Sub

Private Sub CboMonth_GotFocus()
'SendKeys "{F4}"
End Sub

Private Sub EleMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtDay_LostFocus
TxtYear_LostFocus
End Sub
Private Sub TxtDay_Change()
RaiseEvent Change
End Sub

Private Sub TxtDay_GotFocus()
Old_Day = TxtDay.Text
End Sub
Private Sub TxtDay_KeyPress(KeyAscii As Integer)
If IsValidEntry(KeyAscii) = True Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
Private Sub TxtDay_LostFocus()
TxtDay.Text = CStr(Val(TxtDay.Text))
If Val(TxtDay.Text) = 0 Then
    TxtDay.Text = 0 & "1"
ElseIf Val(TxtDay.Text) > 0 And Val(TxtDay.Text) <= 9 Then
     TxtDay.Text = CStr(0 & Val(TxtDay.Text))
ElseIf TxtDay.Text > 30 Then
    TxtDay.Text = Old_Day
End If
End Sub

Private Sub TxtYear_Change()
RaiseEvent Change
End Sub

Private Sub TxtYear_GotFocus()
Old_Year = Trim(TxtYear.Text)
End Sub

Private Sub TxtYear_KeyPress(KeyAscii As Integer)
If IsValidEntry(KeyAscii) = True Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub TxtYear_LostFocus()
If Val(TxtYear.Text) > 2000 Or Val(TxtYear.Text) < 1300 Then
    TxtYear.Text = Old_Year
End If
End Sub

Private Sub UserControl_Initialize()
On Error GoTo ErrTap
Dim StrSetting          As String
Dim StrLastHigriDate    As String
Dim StrLastGergDate     As String
Dim StrLastDay          As String
Dim StrLastMonth        As String
Dim StrLastYear         As String
Dim IntNoDays           As Integer
Dim CalType             As Integer
Dim Today               As Date
With CboMonth
    .AddItem "„Õ—„"
    .itemdata(0) = 0
    .AddItem "’›—"
    .itemdata(1) = 1
    .AddItem "—»Ì⁄ √Ê·"
    .itemdata(2) = 2
    .AddItem "—»Ì⁄  «‰Ï"
    .itemdata(3) = 3
    .AddItem "Ã„«œÏ √Ê·"
    .itemdata(4) = 4
    .AddItem "Ã„«œÏ  «‰Ï"
    .itemdata(5) = 5
    .AddItem "—Ã»"
    .itemdata(6) = 6
    .AddItem "‘⁄»«‰"
    .itemdata(7) = 7
    .AddItem "—„÷«‰"
    .itemdata(8) = 8
    .AddItem "‘Êˆ¯«·"
    .itemdata(9) = 9
    .AddItem "–Ê «·ﬁ⁄œ…"
    .itemdata(10) = 10
    .AddItem "–Ê «·ÕÃ…"
    .itemdata(11) = 11
End With
StrLastGergDate = GetSetting(SystemOptions.SysRegsAppPath, "Higri Date", "Last Date", "")
StrLastHigriDate = GetSetting(SystemOptions.SysRegsAppPath, "Higri Date", "Last Hijri", "")
StrLastDay = GetSetting(SystemOptions.SysRegsAppPath, "Higri Date", "Last Day", "")
StrLastMonth = GetSetting(SystemOptions.SysRegsAppPath, "Higri Date", "Last Month", "")
StrLastYear = GetSetting(SystemOptions.SysRegsAppPath, "Higri Date", "Last Year", "")

If StrLastGergDate <> "" Then
    IntNoDays = Abs(DateDiff("d", CDate(StrLastGergDate), Date))
    If IntNoDays = 0 Then
        Me.Value = StrLastHigriDate
        Exit Sub
    End If
    IntNoDays = IntNoDays + Val(StrLastDay)
    If IntNoDays <= 30 Then
        Me.Value = IntNoDays & "/" & StrLastMonth & "/" & StrLastYear
    Else
        IntNoDays = IntNoDays Mod 30
         Me.Value = IntNoDays & "/" & StrLastMonth & "/" & StrLastYear
    End If
Else
    CurDate
End If
m_BorderColor = EleMain.BackColor
Exit Sub
ErrTap:
CurDate
End Sub
Private Function IsValidEntry(KeyAsc As Integer) As Boolean
If KeyAsc = vbKeyBack Or KeyAsc = vbKeyDelete Then
    IsValidEntry = True
    Exit Function
End If
If InStr(1, "0123456789", Chr(KeyAsc)) <> 0 Then
    IsValidEntry = True
Else
    IsValidEntry = False
End If
End Function

Private Sub UserControl_Resize()
If UserControl.Height >= CboMonth.Height Then
    UserControl.Height = 390
End If
End Sub
Public Property Get Value() As String
Value = m_Value
m_Value = Trim(TxtYear.Text) & "/" & _
IIf((CboMonth.ListIndex + 1) <= 9, "0" & CStr(CboMonth.ListIndex + 1), _
CboMonth.ListIndex + 1) & "/" & Trim(TxtDay.Text)
Value = m_Value
End Property
Public Property Let Value(ByVal vNewValue As String)
On Local Error GoTo ErrTrap
If TypeName(vNewValue) <> "String" Then
    err.Raise 600, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", _
    "ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… «· «—ÌŒ «·„—”· ⁄·Ï ’Ê—… ‰’"
Else
    If vNewValue = "" Then
        Exit Property
    End If
    m_Value = vNewValue
    Dim VarDates As Variant
    VarDates = Split(m_Value, "/", , vbTextCompare)
    If Len(CStr(VarDates(0))) <= 2 Then
        TxtDay.Text = IIf(Val(VarDates(0)) <= 9, "0" & Val(VarDates(0)), VarDates(0))
        TxtYear.Text = VarDates(2)
    Else
        TxtDay.Text = IIf(Val(VarDates(2)) <= 9, "0" & Val(VarDates(2)), VarDates(2))
        TxtYear.Text = VarDates(0)
    End If
    CboMonth.ListIndex = Val(VarDates(1) - 1)
End If
Exit Property
ErrTrap:
    err.Raise 600, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", _
    "ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… «· «—ÌŒ «·„—”· ⁄·Ï ’Ê—… ‰’"
End Property
Public Function Get_Month(Optional IntMonth As Integer = -1) As HijriMonths
If IntMonth = -1 Then
    'Get the Current Month
    Get_Month = CboMonth.ListIndex + 1
End If
End Function
Public Function GetMonthName(Optional IntMonth As Integer = -1) As String
If IntMonth = -1 Then
    GetMonthName = CboMonth.Text
    Exit Function
End If
GetMonthName = choose(IntMonth, "„Õ—„", "’›—", "—»Ì⁄ √Ê·", "—»Ì⁄ À«‰Ï", _
"Ã„«œÏ", "", "", "", "", "")
End Function
Public Property Get DayValue() As Integer
DayValue = Val(TxtDay.Text)
End Property

Public Property Let DayValue(ByVal vNewValue As Integer)
If Val(vNewValue) > 30 Or Val(vNewValue) < 1 Then
    err.Raise 600, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", "«·ﬁÌ„… «·„⁄ÿ«… ··ÌÊ„ €Ì— ’ÕÌÕ…"
Else
    TxtDay.Text = vNewValue
End If
End Property

Public Property Get YearValue() As Integer
YearValue = Val(TxtYear.Text)
End Property

Public Property Let YearValue(ByVal vNewValue As Integer)
If vNewValue > 2000 Or vNewValue < 1000 Then
    err.Raise 600, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï", "«·ﬁÌ„… «·„⁄ÿ«… ··”‰… «·ÂÃ—Ì… €Ì— „ﬁ»Ê·…"
Else
    TxtYear.Text = vNewValue
End If
End Property
Public Property Get BorderColor() As Long
BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal vNewValue As Long)
If Not IsNumeric(vNewValue) Then
    MsgBox "«·ﬁÌ„… «·⁄ÿ«… €Ì— ’ÕÌÕ… ", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "√œ«… «· ﬁÊÌ„ «·ÂÃ—Ï"
Else
    m_BorderColor = vNewValue
    EleMain.BackColor = m_BorderColor
End If
End Property

Private Sub CurDate()
CalType = Calendar
Calendar = vbCalHijri
TxtDay.Text = IIf(Val(day(Date)) <= 9, "0" & Val(day(Date)), day(Date))
TxtYear.Text = Year(Date)
CboMonth.ListIndex = Month(Date) - 1
Calendar = CalType
End Sub
