VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmVacationDays 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «Ì«„ «·⁄ÿ·«  "
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "FrmVacationDays.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   12585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12585
      _cx             =   22199
      _cy             =   14817
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
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   -48
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   -48
         Width           =   12576
         _cx             =   22172
         _cy             =   1032
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "    ”ÃÌ· «Ì«„ «·⁄ÿ·«    "
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
         CaptionStyle    =   1
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
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   336
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   504
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2250
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   7608
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   12360
         _cx             =   21802
         _cy             =   13414
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
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "⁄ÿ·«  «Œ—Ï"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   888
            Left            =   1320
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   372
            Width           =   2412
         End
         Begin MSDataListLib.DataCombo dcVacType 
            Height          =   288
            Left            =   4740
            TabIndex        =   6
            Top             =   504
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDur 
            Height          =   288
            Left            =   8580
            TabIndex        =   7
            Top             =   504
            Width           =   2652
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal FromDateH 
            Height          =   360
            Left            =   8520
            TabIndex        =   8
            Top             =   840
            Width           =   1404
            _ExtentX        =   2487
            _ExtentY        =   635
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   348
            Left            =   9912
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   864
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   101842947
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   360
            Left            =   4680
            TabIndex        =   10
            Top             =   864
            Width           =   1404
            _ExtentX        =   2487
            _ExtentY        =   635
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   348
            Left            =   6120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   864
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   101842947
            CurrentDate     =   41640
         End
         Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
            Height          =   5685
            Left            =   120
            TabIndex        =   12
            Top             =   1755
            Width           =   12180
            _cx             =   21484
            _cy             =   10028
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmVacationDays.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
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
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   372
            Index           =   5
            Left            =   240
            TabIndex        =   13
            Top             =   372
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmVacationDays.frx":0544
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   4
            Left            =   240
            TabIndex        =   14
            Top             =   864
            Width           =   864
            _ExtentX        =   1535
            _ExtentY        =   635
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
            ButtonImage     =   "FrmVacationDays.frx":6DA6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄ÿ·…"
            Height          =   372
            Left            =   7560
            TabIndex        =   19
            Top             =   504
            Width           =   732
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„ «·œ—«”Ï"
            Height          =   372
            Left            =   11280
            TabIndex        =   18
            Top             =   504
            Width           =   972
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   396
            Left            =   11280
            TabIndex        =   17
            Top             =   864
            Width           =   972
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   396
            Left            =   7560
            TabIndex        =   16
            Top             =   864
            Width           =   732
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ"
            Height          =   372
            Left            =   3840
            TabIndex        =   15
            Top             =   504
            Width           =   612
         End
      End
   End
End
Attribute VB_Name = "FrmVacationDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip

Dim rs_Dur  As ADODB.Recordset
Dim rs_vac As ADODB.Recordset
Dim rs_hol As ADODB.Recordset

Dim rsDuration As ADODB.Recordset

Dim FromDate_ As Date
Dim ToDate_ As Date
Dim FromDateH_ As String
Dim ToDateH_ As String



Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.text = "N"
            clear_all Me
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"

        Case 2

          
            Save_WeekVac

        Case 3
          
        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

          '  Del_Company

        Case 5
                save_Vac
        Case 6
            Unload Me
         Case 7
        ' print_report2
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Save_WeekVac()
 
   
    End Sub
    
Private Sub save_Vac()
  Dim BeginTrans As Boolean
  Dim StrSQL As String
  Dim Msg As String
  On Error GoTo errortrap
  
 If dcDur.BoundText = "" Then
 MsgBox ("«Œ — «·⁄«„ «·œ—«”Ï «Ê·«")
 dcDur.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
 If dcVacType.BoundText = "" Then
 MsgBox ("«Œ —‰Ê⁄ «·⁄ÿ·…")
 dcVacType.SetFocus
 SendKeys ("{F4}")
 Exit Sub
 End If
 
  
    Cn.BeginTrans
    BeginTrans = True
    
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT  *  From TblVacationDays "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    rs.AddNew
    rs("id") = CStr(new_id("TblVacationDays", "id", "", True))
    rs("VacationTypeID") = IIf(dcVacType.BoundText = "", Null, val(dcVacType.BoundText))
    rs("VacationType") = IIf(dcVacType.text = "", Null, (dcVacType.text))
    rs("DurationID") = IIf(dcDur.BoundText = "", Null, val(dcDur.BoundText))
    rs("Duration") = IIf(dcDur.text = "", Null, (dcDur.text))
    rs("FromDate") = IIf(IsNull(FromDate.value), Date, FromDate.value)
    rs("ToDate") = IIf(IsNull(ToDate.value), Date, ToDate.value)
    rs("FromDateH") = IIf(IsNull(FromDateH.value), ToHijriDate(Date), FromDateH.value)
    rs("ToDateH") = IIf(IsNull(ToDateH.value), ToHijriDate(Date), ToDateH.value)
    rs("Description") = Text1.text
    rs.update
    
     If FromDate.Enabled = True Then
             
              Add_Schedule FromDate.value, ToDate.value, val(dcDur.BoundText)
     Else
              Add_ScheduleH FromDateH.value, ToDateH.value, val(dcDur.BoundText)
     End If
    
   
    
    Cn.CommitTrans
    BeginTrans = False
    
    
    MsgBox (" „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ")
    Retrive_Vacation (val(dcDur.BoundText))
Exit Sub
errortrap:


    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
   
Private Sub Retrive_Vacation(DurID As Integer)

Dim i As Integer
     Set rs = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT  *  From TblVacationDays where DurationID = " & DurID
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
    FgInstallments.Rows = 1
      
    If rs.RecordCount > 0 Then
         rs.MoveFirst
         
        With FgInstallments
        .Rows = rs.RecordCount + 1
         For i = 1 To FgInstallments.Rows - 1
         .TextMatrix(i, .ColIndex("Serial")) = i
         .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
         .TextMatrix(i, .ColIndex("VacationType")) = IIf(IsNull(rs("VacationType").value), "", rs("VacationType").value)
         .TextMatrix(i, .ColIndex("VacationTypeID")) = IIf(IsNull(rs("VacationTypeID").value), "", rs("VacationTypeID").value)
         .TextMatrix(i, .ColIndex("Duration")) = IIf(IsNull(rs("Duration").value), "", rs("Duration").value)
         .TextMatrix(i, .ColIndex("DurationID")) = IIf(IsNull(rs("DurationID").value), "", rs("DurationID").value)
         .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
         .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
         .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
         .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
         .TextMatrix(i, .ColIndex("Description")) = IIf(IsNull(rs("Description").value), "", rs("Description").value)
         rs.MoveNext
         Next
         End With
    End If

End Sub

   
Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
     
 
 
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    'Dcombos.GetAccountingCodes Me.dcDiscAccount, True

    Dim sql As String
    sql = " select ID , Name  from tbldurations "
    fill_combo dcDur, sql
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
    sql = "  select id , name  from TblVacationTypes  "
    Else
    sql = " select id , namee  from TblVacationTypes "
    End If
    fill_combo dcVacType, sql
       
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «‰Ê«⁄ «·„Œ«·›«   "
    LogTextE = " Open Window " & "  Violation Types "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    Dim My_SQL As String
       
          
    Resize_Form Me
      
    Me.TxtModFlg.text = "R"
    ' XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

FromDate.value = Date
ToDate.value = Date

    Exit Sub

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
 '   Dim IntResult As String
 '   Dim StrMSG As String
 ''   On Error GoTo ErrTrap
'
'    If Me.TxtModFlg.text <> "R" Then
'
'        Select Case Me.TxtModFlg.text
''
'            Case "N"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save " & Chr(13)
'                    StrMSG = StrMSG & " the new data  " & Chr(13)
'                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
'                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
'                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
'
'                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
'                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
'                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
'                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
'
'                End If
'
''            Case "E"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save  " & Chr(13)
'                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
'                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
''                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
'                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
'
'                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
'                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
'                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
'                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
'
'                End If
'
'        End Select
'
'        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
'
'        Select Case IntResult
'
'            Case vbYes
'                Cancel = True
'
'                'SaveData
'
'            Case vbCancel
'                Cancel = True
'        End Select
'
'    End If
'
'    Exit Sub
'ErrTrap:
End Sub
'
Private Sub ChangeLang()
 
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «‰Ê«⁄ «·„Œ«·›«   "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub FromDate_Change()
     
    If FromDate.value > ToDate.value Then
            ToDate.value = FromDate.value
    End If
    
    If FromDate.value < FromDate_ Or FromDate.value > ToDate_ Then
            FromDate.value = FromDate_
    End If
    
    If Me.TxtModFlg.text <> "R" Then
        FromDateH.value = ToHijriDate(FromDate.value)
    End If
        
End Sub

Private Sub FromDateH_LostFocus()

  If FromDateH.value < FromDateH_ Or FromDateH.value > ToDateH_ Then
            FromDateH.value = FromDateH_
    End If


      VBA.Calendar = vbCalGreg
       FromDate.value = ToGregorianDate(FromDateH.value)
      
End Sub

Private Sub ToDate_Change()
  If ToDate.value < FromDate_ Or ToDate.value > ToDate_ Then
            ToDate.value = ToDate_
    End If

   If Me.TxtModFlg.text <> "R" Then
        ToDateH.value = ToHijriDate(ToDate.value)
     End If
End Sub

Private Sub ToDateH_LostFocus()

    If ToDateH.value < FromDateH_ Or ToDateH.value > ToDateH_ Then
    
            ToDateH.value = ToDateH_
    End If
    
   VBA.Calendar = vbCalGreg
   ToDate.value = ToGregorianDate(ToDateH.value)
        
End Sub


Private Sub Del_row()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
        Dim str As String, sr As String
        str = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("id"))
        sr = FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("serial"))
        
        If str <> "" Then
 
        Msg = "”Ì „ Õ–› »Ì«‰«  ”ÿ— —ﬁ„ " & Chr(13)
        Msg = Msg + (sr) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblVacationDays  where  ID =" & val(str)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "SELECT  *  From TblViolationTypes"
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          
                If rs.RecordCount < 1 Then
                    'clear_all Me
                    'TxtModFlg_Change
                    'XPTxtCurrent.Caption = 0
                    'XPTxtCount.Caption = 0
                Else
                   Retrive_Vacation (val(dcDur.BoundText))
                End If
            End If
        End If

    Else
        'clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       ' TxtModFlg_Change
        Exit Sub
    End If
 Retrive_Vacation (val(dcDur.BoundText))
    'TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Œ«·›… "
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Add_Schedule(FromDate As Date, ToDate As Date, dur As Integer)
  
   Dim str As String, str1 As String
   Do While FromDate <= ToDate
        str = Weekday(FromDate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If IsHoliday(str1) Or ISOfficialVacation(FromDate, dur) Then
                 AddRowToSchedule dur, FromDate, ToHijriDate(FromDate), True
        Else
                 AddRowToSchedule dur, FromDate, ToHijriDate(FromDate), False
        End If
        VBA.Calendar = vbCalGreg
       FromDate = DateAdd("d", 1, FromDate)
   Loop
End Sub


Private Sub Add_ScheduleH(FromDate As String, ToDate As String, dur As Integer)
   Dim str As String, str1 As String
   VBA.Calendar = vbCalHijri
   
   
  FromDate = Format(FromDate, "yyyy/MM/dd")
  ToDate = Format(ToDate, "yyyy/MM/dd")
  
   
   Do While FromDate <= ToDate
        str = Weekday(FromDate, vbSaturday)
        str1 = WeekdayName(str, True, vbSaturday)
        
        If ISOfficialVacationH(FromDate, dur) Then
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, True
        Else
                 AddRowToSchedule dur, ToGregorianDate(FromDate), FromDate, False
        End If
        VBA.Calendar = vbCalHijri
        FromDate = DateAdd("d", 1, FromDate)
         FromDate = Format(FromDate, "yyyy/MM/dd")
        VBA.Calendar = vbCalGreg
   Loop
   VBA.Calendar = vbCalGreg
End Sub



Private Function IsHoliday(day As String) As Boolean


    Dim str As String
    str = " select * from  tblholidays  "
    Set rs_hol = New ADODB.Recordset
    rs_hol.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs_hol.RecordCount > 0 Then
            If rs_hol("sa").value = True And day = "Sat" Then
                IsHoliday = True
            ElseIf rs_hol("su").value = True And day = "Sun" Then
                IsHoliday = True
            ElseIf rs_hol("Mo").value = True And day = "Mon" Then
                IsHoliday = True
            ElseIf rs_hol("Tu").value = True And day = "Tue" Then
                 IsHoliday = True
            ElseIf rs_hol("We").value = True And day = "Wed" Then
                IsHoliday = True
            ElseIf rs_hol("Th").value = True And day = "Thu" Then
                IsHoliday = True
            ElseIf rs_hol("Fr").value = True And day = "Fri" Then
                IsHoliday = True
            End If
    End If

End Function

Private Function ISOfficialVacation(dt As Date, dur As Integer) As Boolean
    
    Dim str As String
    str = " select * from  TblVacationDays  where DurationID =   " & dur & "  and   '" & dt & "' Between FromDate And ToDate  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacation = True
    End If

End Function

Private Function ISOfficialVacationH(dt As String, dur As Integer) As Boolean
    
    Dim str As String
    str = " select * from  TblVacationDays  where DurationID =   " & dur & "  and    '" & dt & "'   >=  FromDateH  And   '" & dt & "'  <=   ToDateH  "
    Set rs_vac = New ADODB.Recordset
    rs_vac.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If rs_vac.RecordCount > 0 Then
        ISOfficialVacationH = True
    End If

End Function



Private Sub AddRowToSchedule(dur As Integer, dt As Date, dth As String, isvac As Boolean)
       
       Dim str As String
       If FromDate.Enabled = True Then
                str = " Select   *  from  TblVacationschedule where Date = '" & dt & "' and  DurationID = " & dur
       Else
                str = " Select   *  from  TblVacationschedule where DateH = '" & Format(dth, "yyyy/MM/dd") & "' and DurationID = " & dur
       End If
       Set rs_Dur = New ADODB.Recordset
       rs_Dur.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
       
       If rs_Dur.RecordCount > 0 Then
            rs_Dur.MoveFirst
            rs_Dur("isvac") = isvac
            rs_Dur.update
       End If

End Sub


