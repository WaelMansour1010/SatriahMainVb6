VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.UserControl NourCalendar 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   ScaleHeight     =   2880
   ScaleWidth      =   4335
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   2880
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4335
      _cx             =   7646
      _cy             =   5080
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
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
      ResizeFonts     =   -1  'True
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"NourCalendar.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.ComboBox CboYears 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   15
         Width           =   1065
      End
      Begin VB.ComboBox CboMonth 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3255
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   15
         Width           =   1065
      End
      Begin VSFlex8LCtl.VSFlexGrid Fg 
         Height          =   2130
         Left            =   15
         TabIndex        =   4
         Top             =   735
         Width           =   4305
         _cx             =   7594
         _cy             =   3757
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12582912
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   3
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   3
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Nour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   2145
      End
   End
End
Attribute VB_Name = "NourCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event DateClicked(C_Date As Date)

Public Event DateDblClick(DateClicked As Date)

Public Event DateChanged()

Public Event DateChecked(C_Date As Date, State As CheckSate)

Public Event CalendarMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event CalendarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event CalendarMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event HeaderMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event HeaderMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event HeaderMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event HeaderClicked()

Public Event HeaderDblClick()

Public Event DateMoveOver(M_Date As Date)

Public Event DateMouseUp(M_Date As Date, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim m_OldYear  As String
Dim m_OldMonth As String

Private m_CurrentMonth As Integer

Private m_CurrentYear As Integer

Private m_ShowWeekNumber As Boolean

Private m_EnableMonthList As Boolean

Private m_EnableYearList As Boolean

Private m_CalendarToolTip As String

Dim m_ShowCheckBox As Boolean
Dim m_ColDateVar As New Collection

'Dim m_ColDateImage As New Collection

Public Enum CheckSate
    DateUnChecked
    DateChecked
End Enum

Private Type Row_Col
    Row As Long
    Col As Long
End Type

Private Sub CboMonth_Change()
    RaiseEvent DateChanged
End Sub

Private Sub CboMonth_Click()
    DisDates
    RaiseEvent DateChanged
End Sub

Private Sub CboYears_Change()
    RaiseEvent DateChanged
End Sub

Private Sub CboYears_Click()
    DisDates
    RaiseEvent DateChanged
End Sub

Private Sub EleMain_ResizeChildren()
    LayoutFg
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With FG

        If FG.Cell(flexcpChecked, Row, Col) <> flexNoCheckbox Then
            If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                RaiseEvent DateChecked(GetDate(Row, Col), DateChecked)
            Else
                RaiseEvent DateChecked(GetDate(Row, Col), DateUnChecked)
            End If
        End If

    End With

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    With FG

        If FG.TextMatrix(Row, Col) = "" Then
            Cancel = True
        End If

    End With

End Sub

Private Sub fg_Click()
    Dim temp As String

    If FG.Row = -1 Then Exit Sub
    If FG.Col = -1 Then Exit Sub
    If FG.Row < FG.FixedRows Then
        Exit Sub
    End If

    If FG.Col = FG.ColIndex("WeekNumber") Then
    
    Else

        If FG.TextMatrix(FG.Row, FG.Col) <> "" Then
            temp = GetDate(FG.Row, FG.Col)
            RaiseEvent DateClicked(CDate(temp))
        End If
    End If

End Sub

Private Sub Fg_DblClick()

    Dim temp As String

    If FG.Row = -1 Then Exit Sub
    If FG.Col = -1 Then Exit Sub
    If FG.Row < FG.FixedRows Then
        Exit Sub
    End If

    If FG.Col = FG.ColIndex("WeekNumber") Then
    
    Else

        If FG.TextMatrix(FG.Row, FG.Col) <> "" Then
            temp = GetDate(FG.Row, FG.Col)
            RaiseEvent DateDblClick(CDate(temp))
        End If
    End If

End Sub

Private Sub Fg_MouseDown(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    RaiseEvent CalendarMouseDown(Button, Shift, X, Y)
End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long
    Dim M_Date As Date

    With FG
        LngMouseRow = .MouseRow
        LngMouseCol = .MouseCol
    End With

    If LngMouseCol <= -1 Or LngMouseRow <= -1 Then
        RaiseEvent CalendarMouseMove(Button, Shift, X, Y)
    ElseIf Trim(FG.TextMatrix(LngMouseRow, LngMouseCol)) = "" Then
        RaiseEvent CalendarMouseMove(Button, Shift, X, Y)

    ElseIf LngMouseRow >= FG.FixedRows And LngMouseCol > FG.ColIndex("WeekNumber") Then
        M_Date = CDate(GetDate(LngMouseRow, LngMouseCol))

        If IsDate(M_Date) Then
            RaiseEvent DateMoveOver(M_Date)
        End If

    Else
        RaiseEvent CalendarMouseMove(Button, Shift, X, Y)
    End If

End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long
    Dim M_Date As Date

    With FG
        LngMouseRow = .MouseRow
        LngMouseCol = .MouseCol
    End With

    If LngMouseCol <= -1 Or LngMouseRow <= -1 Then
        RaiseEvent CalendarMouseUp(Button, Shift, X, Y)
    ElseIf Trim(FG.TextMatrix(LngMouseRow, LngMouseCol)) = "" Then
        RaiseEvent CalendarMouseUp(Button, Shift, X, Y)

    ElseIf LngMouseRow >= FG.FixedRows And LngMouseCol > FG.ColIndex("WeekNumber") Then
        M_Date = CDate(GetDate(LngMouseRow, LngMouseCol))

        If IsDate(M_Date) Then
            RaiseEvent DateMouseUp(M_Date, Button, Shift, X, Y)
        End If

    Else
        RaiseEvent CalendarMouseUp(Button, Shift, X, Y)
    End If

End Sub

Private Sub lblCap_Click()
    RaiseEvent HeaderClicked
End Sub

Private Sub lblCap_DblClick()
    RaiseEvent HeaderDblClick
End Sub

Private Sub lblCap_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    RaiseEvent HeaderMouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCap_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    RaiseEvent HeaderMouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblCap_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    RaiseEvent HeaderMouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer

    If UserControl.RightToLeft = True Then
        CboMonth.RightToLeft = True
        CboYears.RightToLeft = True

        With CboMonth
            .Clear
            .AddItem "íäÇíÑ"
            .ItemData(0) = 1
            .AddItem "ÝÈÑÇíÑ"
            .ItemData(1) = 2
            .AddItem "ãÇÑÓ"
            .ItemData(2) = 3
            .AddItem "ÇÈÑíá"
            .ItemData(3) = 4
            .AddItem "ãÇíæ"
            .ItemData(4) = 5
            .AddItem "íæäíæ"
            .ItemData(5) = 6
            .AddItem "íæáíæ"
            .ItemData(6) = 7
            .AddItem "ÃÛÓØÓ"
            .ItemData(7) = 8
            .AddItem "ÓÈÊãÈÑ"
            .ItemData(8) = 9
            .AddItem "ÃßÊæÈÑ"
            .ItemData(9) = 10
            .AddItem "äæÝãÈÑ"
            .ItemData(10) = 11
            .AddItem "ÏíÓãÈÑ"
            .ItemData(11) = 12
            .ListIndex = Month(Date) - 1
        End With

        lblCap.RightToLeft = True
        'lblCap.Alignment = vbLeftJustify
    
    Else
        CboMonth.RightToLeft = False
        CboYears.RightToLeft = False

        With CboMonth
        
            .Clear
            .AddItem "January"
            .ItemData(0) = 1
            .AddItem "February"
            .ItemData(1) = 2
            .AddItem "March"
            .ItemData(2) = 3
            .AddItem "April"
            .ItemData(3) = 4
            .AddItem "May"
            .ItemData(4) = 5
            .AddItem "June"
            .ItemData(5) = 6
            .AddItem "July"
            .ItemData(6) = 7
            .AddItem "August"
            .ItemData(7) = 8
            .AddItem "September"
            .ItemData(8) = 9
            .AddItem "October"
            .ItemData(9) = 10
            .AddItem "November"
            .ItemData(10) = 11
            .AddItem "December"
            .ItemData(11) = 12
            .ListIndex = Month(Date) - 1
        End With

        lblCap.RightToLeft = False
        'lblCap.Alignment = vbLeftJustify
    
    End If

    With CboYears
        .Clear

        For i = 1900 To 2100
            .AddItem i
        Next i

        .ListIndex = year(Date) - 1900
    End With

    LoadFgSetting
    lblCap.Caption = ""
    DisDates
End Sub

Private Sub DisDates(Optional BolShowWeekNumber As Boolean = True)
    Dim i As Integer
    Dim MaxDay As Integer
    Dim CurMonth As Integer
    Dim PutDate As Date
    Dim CurYear As String
    Dim CurDay As Integer
    Dim X As Long
    Dim Y As Long
    Dim Ipic As IPictureDisp
    Dim TempPic As Variant
    Dim TempVar As Variant
    Dim k As Integer

    If CboMonth.ListIndex = -1 Then Exit Sub
    If CboYears.ListIndex = -1 Then Exit Sub
    If val(CboYears.Text) = 0 Then Exit Sub

    CurMonth = CboMonth.ItemData(CboMonth.ListIndex)
    CurYear = CboYears.List(CboYears.ListIndex)
    lblCap.Caption = CboMonth.Text & " " & CurYear

    Select Case CboMonth.ItemData(CboMonth.ListIndex)

        Case 1, 3, 5, 7, 8, 10, 12
            MaxDay = 31

        Case 2

            If Month(DateAdd("d", 1, CDate("28/2/" & CboYears.List(CboYears.ListIndex)))) = 2 Then
                MaxDay = 29
            Else
                MaxDay = 28
            End If

        Case Else
            MaxDay = 30
    End Select

    X = Weekday(CDate("01/0" & CurMonth & "/" & CurYear & ""), vbSaturday)
    X = X
    Y = FG.FixedRows
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Cell(flexcpChecked, 1, 0, FG.Rows - 1, FG.Cols - 1) = flexNoCheckbox
    LoadFgSetting

    For i = 1 To MaxDay

        FG.TextMatrix(Y, X) = i
    
        If m_ShowCheckBox = True Then
            FG.Cell(flexcpChecked, Y, X) = flexUnchecked
        End If

        If CurMonth = Month(Date) And CurYear = year(Date) Then
            If day(Date) = i Then
                FG.Cell(flexcpForeColor, Y, X) = vbRed
                FG.TextMatrix(Y, X) = i & CHR(13) & "Çáíæã"
            End If
        End If

        PutDate = CDate(CurYear & "/" & CurMonth & "/" & i)

        If BolShowWeekNumber = True Then
            FG.TextMatrix(Y, FG.ColIndex("WeekNumber")) = GetTheWeek(PutDate)
        End If

        '    If ImgList.ListImages.Count > 0 Then
        '        For k = 1 To ImgList.ListImages.Count
        '            If CStr(PutDate) = ImgList.ListImages(k).Key Then
        '                Fg.Cell(flexcpPicture, Y, X) = ImgList.ListImages(k).ExtractIcon
        '                Debug.Print Y & "--" & X & "Photo"
        '                Exit For
        '            Else
        '                Fg.Cell(flexcpPicture, Y, X) = Nothing
        '                Debug.Print Y & "--" & X & " No Photo"
        '            End If
        '        Next k
        '    End If
    
        If X >= 7 Then
            X = 1
            Y = Y + 1
        Else
            X = X + 1
        End If

    Next i

    FG.Redraw = flexRDBuffered
    FG.Refresh

End Sub

Public Function GetTheWeek(TheDate As Date) As Integer

    Const StartDayMonth As String = "1/1/"
    Dim YearStartDate As Date

    If Month(TheDate) = 12 And day(TheDate) > 25 And Weekday(TheDate, vbSaturday) < Weekday(StartDayMonth & year(TheDate) + 1, vbSaturday) Then

        YearStartDate = StartDayMonth & year(TheDate) + 1
    Else
        YearStartDate = StartDayMonth & year(TheDate)
    End If

    GetTheWeek = (TheDate - YearStartDate + Weekday(YearStartDate, vbSaturday) - 1) \ 7 + 1

End Function

Private Function GetDate(Row As Long, _
                         Col As Long) As Date
    Dim temp As String
    GetDate = Date

    With FG

        If .TextMatrix(Row, Col) = "" Then
            GetDate = GetDate
            Exit Function
        End If

        temp = FG.TextMatrix(Row, Col) & "/" & CboMonth.ItemData(CboMonth.ListIndex) & "/" & CboYears.List(CboYears.ListIndex) & ""
        temp = Replace(temp, "Çáíæã", "", , , vbTextCompare)
        GetDate = CDate(temp)
    End With

End Function

Public Function CheckDate(C_Date As Date) As Boolean

End Function

Public Function SetDateVariant(C_Date As Date, _
                               VarData As Variant) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim IntDay As Integer

    If IsEmpty(VarData) Then
        SetDateVariant = False
        Exit Function
    End If

    If Me.CurrentMonth <> Month(C_Date) Then
        SetDateVariant = False
        Exit Function
    End If

    If Me.CurrentYear <> year(C_Date) Then
        SetDateVariant = False
        Exit Function
    End If

    IntDay = day(C_Date)

    With FG

        For i = .FixedRows To .Rows - 1
            For j = .ColIndex("Saturday") To .Cols - 1

                If val(.TextMatrix(i, j)) = IntDay Then
                    .Cell(flexcpData, i, j) = VarData
                    SetDateVariant = True
                    Exit Function
                End If

            Next j
        Next i

    End With

End Function

Public Property Get DateVariant(C_Date As Date) As Variant
    Dim temp As Variant
    Dim i As Integer

    For i = 1 To m_ColDateVar.count
        temp = m_ColDateVar.Item(i)

        If temp(1) = C_Date Then
            DateVariant = temp(2)
        End If

    Next i

End Property

Public Property Let DateVariant(C_Date As Date, ByVal vNewValue As Variant)
    Dim temp As Variant
    ReDim temp(3) As Variant
    temp(0) = ""
    temp(1) = C_Date
    temp(2) = vNewValue
    m_ColDateVar.Add temp
End Property

Private Sub UserControl_Terminate()
    Set m_ColDateVar = Nothing
End Sub

'Public Property Get DateImage(C_Date As Date) As IPictureDisp
'Dim I As Integer
'For I = 1 To ImgList.ListImages.Count
'    If CStr(C_Date) = ImgList.ListImages(I).Key Then
'        Set DateImage = ImgList.ListImages(I).Picture
'    End If
'Next I
'End Property
'Public Property Set DateImage(C_Date As Date, ByVal vNewValue As IPictureDisp)
'If Me.CurrentMonth <> Month(C_Date) Then
'    Exit Property
'End If
'If Me.CurrentYear <> Year(C_Date) Then
'    Exit Property
'End If
'With Fg
'
'End With
'Dim I As Integer
'I = 1
'Do While I <= ImgList.ListImages.Count
'    If ImgList.ListImages(I).Key = CStr(C_Date) Then
'        ImgList.ListImages.Remove I
'        I = I - 1
'    End If
'    I = I + 1
'Loop
'ImgList.ListImages.Add , CStr(C_Date), vNewValue
'DisDates
'End Property
Public Property Get ShowCheckBox() As Boolean
    ShowCheckBox = m_ShowCheckBox
End Property

Public Property Let ShowCheckBox(ByVal vNewValue As Boolean)
    m_ShowCheckBox = vNewValue
    DisDates
End Property

Private Sub LoadFgSetting()
    Dim i As Integer

    With FG
        .AllowSelection = True
        .AllowBigSelection = False
        .SelectionMode = flexSelectionFree
        .BackColorSel = FG.backcolor
        .ForeColorSel = .ForeColor
        .GridLines = flexGridFlat
        .Cell(flexcpForeColor, .FixedRows, .ColIndex("WeekNumber"), .Rows - 1, .ColIndex("WeekNumber")) = &HC0&
        .Cell(flexcpFontName, .FixedRows, .ColIndex("WeekNumber"), .Rows - 1, .ColIndex("WeekNumber")) = "Tahoma"
        .Cell(flexcpFontBold, .FixedRows, .ColIndex("WeekNumber"), .Rows - 1, .ColIndex("WeekNumber")) = True
        .Cell(flexcpFontSize, .FixedRows, .ColIndex("WeekNumber"), .Rows - 1, .ColIndex("WeekNumber")) = 7
        .Cell(flexcpBackColor, .FixedRows, .ColIndex("WeekNumber"), .Rows - 1, .ColIndex("WeekNumber")) = vbYellow

        If UserControl.RightToLeft = True Then
            FG.RightToLeft = True

            For i = FG.ColIndex("Saturday") To FG.Cols - 1
                FG.ColAlignment(i) = flexAlignRightCenter
                FG.FixedAlignment(i) = flexAlignRightCenter
            Next i

            FG.ColAlignment(FG.ColIndex("WeekNumber")) = flexAlignRightBottom
            FG.TextMatrix(0, .ColIndex("Saturday")) = "ÇáÓÈÊ"
            FG.TextMatrix(0, .ColIndex("Sunday")) = "ÇáÃÍÏ"
            FG.TextMatrix(0, .ColIndex("Monday")) = "ÇáÃËäíä"
            FG.TextMatrix(0, .ColIndex("Tuesday")) = "ÇáËáÇËÇÁ"
            FG.TextMatrix(0, .ColIndex("Wednesday")) = "ÇáÃÑÈÚÇÁ"
            FG.TextMatrix(0, .ColIndex("Thursday")) = "ÇáÎãíÓ"
            FG.TextMatrix(0, .ColIndex("Friday")) = "ÇáÌãÚÉ"
        ElseIf UserControl.RightToLeft = False Then
            FG.RightToLeft = Fals

            For i = FG.ColIndex("Saturday") To FG.Cols - 1
                FG.ColAlignment(i) = flexAlignLeftCenter
                FG.FixedAlignment(i) = flexAlignLeftCenter
            Next i

            FG.ColAlignment(FG.ColIndex("WeekNumber")) = flexAlignLeftBottom
            FG.TextMatrix(0, .ColIndex("Saturday")) = "Sat"
            FG.TextMatrix(0, .ColIndex("Sunday")) = "Sun"
            FG.TextMatrix(0, .ColIndex("Monday")) = "Mon"
            FG.TextMatrix(0, .ColIndex("Tuesday")) = "Tues"
            FG.TextMatrix(0, .ColIndex("Wednesday")) = "Wednes"
            FG.TextMatrix(0, .ColIndex("Thursday")) = "Thurs"
            FG.TextMatrix(0, .ColIndex("Friday")) = "Fri"
        End If
    
    End With

End Sub

Public Property Get GetMonthDays() As Integer
    Dim MaxDay As Integer

    Select Case CboMonth.ItemData(CboMonth.ListIndex)

        Case 1, 3, 5, 7, 8, 10, 12
            MaxDay = 31

        Case 2

            If Month(DateAdd("d", 1, CDate("28/2/" & CboYears.List(CboYears.ListIndex)))) = 2 Then
                MaxDay = 29
            Else
                MaxDay = 28
            End If

        Case Else
            MaxDay = 30
    End Select

    GetMonthDays = MaxDay
End Property

Public Property Get CurrentMonth() As Integer
    m_CurrentMonth = CboMonth.ItemData(CboMonth.ListIndex)
    CurrentMonth = m_CurrentMonth

End Property

Public Property Let CurrentMonth(ByVal vNewValue As Integer)

    If vNewValue <= 0 Or vNewValue > 12 Then
        Err.Raise 1005, "NourCalendar:CurrentMonth", "Invalid Value"
    Else
        m_CurrentMonth = vNewValue - 1
        CboMonth.ListIndex = m_CurrentMonth
        DisDates
        RaiseEvent DateChanged
    End If

End Property

Public Function GetMonthDayCount(m_Day As VbDayOfWeek) As Integer
    Dim i As Integer
    Dim IntCount  As Integer
    Dim IntColIndex As Integer

    IntColIndex = GetDayIndex(m_Day)

    For i = FG.FixedRows To FG.Rows - 1

        If FG.TextMatrix(i, IntColIndex) <> "" Then
            IntCount = IntCount + 1
        End If

    Next i

    GetMonthDayCount = IntCount
End Function

Public Property Get ShowWeekNumber() As Boolean
    ShowWeekNumber = m_ShowWeekNumber
End Property

Public Property Let ShowWeekNumber(ByVal vNewValue As Boolean)
    m_ShowWeekNumber = vNewValue

    If m_ShowWeekNumber = True Then
        FG.ColHidden(FG.ColIndex("WeekNumber")) = False
    Else
        FG.ColHidden(FG.ColIndex("WeekNumber")) = True
    End If

    LayoutFg
End Property

Private Sub LayoutFg()
    Dim s As Single
    Dim i As Integer

    'Fg.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    With FG

        If m_ShowWeekNumber = True Then
            s = .Width / .Cols
            .ColWidth(.ColIndex("WeekNumber")) = (s * (2 / 3))
            s = (.Width - .ColWidth(.ColIndex("WeekNumber"))) / (.Cols - 1)
        Else
            s = .Width / (.Cols - 1)
        End If
    
        For i = 1 To .Cols - 1
            .ColWidth(i) = s
        Next i

        s = .ClientHeight / 7
        .RowHeightMin = s
        .Redraw = flexRDDirect
    End With

End Sub

Public Property Get CurrentYear() As Integer
    m_CurrentYear = val(CboYears.Text)
    CurrentYear = m_CurrentYear
End Property

Public Property Let CurrentYear(ByVal vNewValue As Integer)

    If vNewValue <= 1899 Or vNewValue > 2100 Then
        Err.Raise 1005, "NourCalendar:CurrentYear", "Invalid Value"
    Else
        m_CurrentYear = vNewValue
        CboYears.Text = m_CurrentYear
        DisDates
        RaiseEvent DateChanged
    End If

End Property

Public Property Get EnableMonthList() As Boolean
    EnableMonthList = m_EnableMonthList
End Property

Public Property Let EnableMonthList(ByVal vNewValue As Boolean)
    m_EnableMonthList = vNewValue

    If m_EnableMonthList = True Then
        CboMonth.Enabled = True
    Else
        CboMonth.Enabled = False
    End If

End Property

Public Property Get EnableYearList() As Boolean
    EnableYearList = m_EnableYearList
End Property

Public Property Let EnableYearList(ByVal vNewValue As Boolean)
    m_EnableYearList = vNewValue

    If m_EnableYearList = True Then
        CboYears.Enabled = True
    Else
        CboYears.Enabled = False
    End If

End Property

Public Sub SetDayImage(m_Day As VbDayOfWeek)
    Dim IntColIndex As Integer
    IntColIndex = GetDayIndex(m_Day)

End Sub

Private Function GetDayIndex(m_Day As VbDayOfWeek) As Integer
    Dim IntColIndex As Integer

    If m_Day = vbSaturday Then
        IntColIndex = 1
    ElseIf m_Day = vbSunday Then
        IntColIndex = 2
    ElseIf m_Day = vbMonday Then
        IntColIndex = 3
    ElseIf m_Day = vbTuesday Then
        IntColIndex = 4
    ElseIf m_Day = vbWednesday Then
        IntColIndex = 5
    ElseIf m_Day = vbThursday Then
        IntColIndex = 6
    ElseIf m_Day = vbFriday Then
        IntColIndex = 7
    End If

    GetDayIndex = IntColIndex
End Function

Public Sub ClearCalendarImages()
    Dim i  As Integer
    Dim j As Integer

    With FG

        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1

                If Not .Cell(flexcpPicture, i, j) Is Nothing Then
                    Set .Cell(flexcpPicture, i, j) = Nothing
                End If

            Next
        Next i

    End With

End Sub

Public Function SetDateImage(C_Date As Date, _
                             XPic As IPictureDisp) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim IntDay As Integer

    If XPic Is Nothing Then
        SetDateImage = False
        Exit Function
    End If

    If Me.CurrentMonth <> Month(C_Date) Then
        SetDateImage = False
        Exit Function
    End If

    If Me.CurrentYear <> year(C_Date) Then
        SetDateImage = False
        Exit Function
    End If

    IntDay = day(C_Date)

    With FG

        For i = .FixedRows To .Rows - 1
            For j = .ColIndex("Saturday") To .Cols - 1

                If val(.TextMatrix(i, j)) = IntDay Then
                    Set .Cell(flexcpPicture, i, j) = XPic
                    SetDateImage = True
                    Exit Function
                End If

            Next j
        Next i

    End With

End Function

Public Function GetDateImage(C_Date As Date) As IPictureDisp
    Dim i As Integer
    Dim j As Integer
    Dim IntDay As Integer

    If Me.CurrentMonth <> Month(C_Date) Then
        Set GetDateImage = Nothing
        Exit Function
    End If

    If Me.CurrentYear <> year(C_Date) Then
        Set GetDateImage = Nothing
        Exit Function
    End If

    IntDay = day(C_Date)

    With FG

        For i = .FixedRows To .Rows - 1
            For j = .ColIndex("Saturday") To .Cols - 1

                If val(.TextMatrix(i, j)) = IntDay Then
                    Set GetDateImage = .Cell(flexcpPicture, i, j)
                    Exit Function
                End If

            Next j
        Next i

    End With

End Function

Public Property Get CalendarToolTip() As String
    CalendarToolTip = m_CalendarToolTip
End Property

Public Property Let CalendarToolTip(ByVal vNewValue As String)
    m_CalendarToolTip = vNewValue
    FG.ToolTipText = m_CalendarToolTip
End Property

Public Function GetDateVariant(C_Date As Date) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim IntDay As Integer

    If Me.CurrentMonth <> Month(C_Date) Then
        Set GetDateVariant = Nothing
        Exit Function
    End If

    If Me.CurrentYear <> year(C_Date) Then
        Set GetDateVariant = Nothing
        Exit Function
    End If

    IntDay = day(C_Date)

    With FG

        For i = .FixedRows To .Rows - 1
            For j = .ColIndex("Saturday") To .Cols - 1

                If val(.TextMatrix(i, j)) = IntDay Then
                    GetDateVariant = .Cell(flexcpData, i, j)
                    Exit Function
                End If

            Next j
        Next i

    End With

End Function
