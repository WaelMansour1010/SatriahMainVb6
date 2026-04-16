VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaysHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì· «Ì«„ «·”‰Ê«  «·œ—«”Ì…"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   Icon            =   "frmDaysHistory.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   10965
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
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10965
      _cx             =   19341
      _cy             =   15584
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   720
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   645
         Width           =   10965
         _cx             =   19341
         _cy             =   1270
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
         Align           =   0
         AutoSizeChildren=   0
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
         Begin MSDataListLib.DataCombo DcYear 
            Height          =   288
            Left            =   6720
            TabIndex        =   4
            Top             =   240
            Width           =   3132
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   264
            Left            =   3888
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   1524
            _ExtentX        =   2699
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98959363
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal FromdateH 
            Height          =   264
            Left            =   3828
            TabIndex        =   7
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   324
            Left            =   672
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   1524
            _ExtentX        =   2699
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98959363
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal ToDateH 
            Height          =   324
            Left            =   720
            TabIndex        =   9
            Top             =   240
            Width           =   1392
            _ExtentX        =   2461
            _ExtentY        =   582
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì‰ ÂÏ ›Ï"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   1944
            TabIndex        =   11
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì»œ√ „‰ "
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   5340
            TabIndex        =   10
            Top             =   240
            Width           =   948
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   324
            Index           =   0
            Left            =   9744
            TabIndex        =   5
            Top             =   240
            Width           =   1104
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   660
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10995
         _cx             =   19394
         _cy             =   1164
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
         Caption         =   " ›’Ì· «Ì«„ «·”‰Ê«   "
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
      Begin VSFlex8UCtl.VSFlexGrid fg 
         Height          =   7230
         Left            =   0
         TabIndex        =   2
         Top             =   1410
         Width           =   11025
         _cx             =   19447
         _cy             =   12753
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDaysHistory.frx":000C
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
   End
End
Attribute VB_Name = "frmDaysHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip

'Public DurID As Integer

Private Sub ChangeLang()
   
   
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
  
  
  
End Sub


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  »‰ﬂ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·»‰ﬂ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·»‰ﬂ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
      '  .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·»‰ﬂ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ »‰ﬂ" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
     '   .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
   '     .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «·»‰Êﬂ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dim str As String
    str = " select id ,name from tbldurations "
    fill_combo DcYear, str

    ScreenNameArabic = "»Ì«‰«  «·‘Ì› "
    ScreenNameEnglish = " Shift Data "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    Resize_Form Me
   
  'Retrive_Det
    
         
   AddTip
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive_Det2(DurID As Integer, Typ As Integer)

 Set rs = New ADODB.Recordset
    Dim str As String
    If Typ = -1 Then Exit Sub
    If Typ = 0 Then
'    str = "   SELECT  TblDurations2.id YID , dbo.TblVacationschedule22.Day,  TblVacationschedule22.color , dbo.TblVacationschedule22.DurationID, dbo.TblDurations_Details2.Name, dbo.TblDurations_Details2.ID, dbo.TblVacationschedule22.Date   , dbo.TblVacationschedule22.DateH   , dbo.TblVacationschedule22.ISVac, "
'    str = str & "  dbo.TblDurations_Details2.FromDate , dbo.TblDurations_Details2.ToDate"
'    str = str & "  FROM     dbo.TblDurations2 INNER JOIN"
'    str = str & "  dbo.TblDurations_Details2 ON dbo.TblDurations2.ID = dbo.TblDurations_Details2.DID INNER JOIN"
'    str = str & "  dbo.TblVacationschedule22 ON dbo.TblDurations2.ID = dbo.TblVacationschedule22.DurationID"
'    str = str & "  and  (dbo.TblVacationschedule22.Date BETWEEN dbo.TblDurations_Details2.FromDate AND dbo.TblDurations_Details2.ToDate)  "
'    str = str & "  Where 1=1   "
'
    str = "SELECT     dbo.Tbldurations2.ID AS YID, dbo.TblVacationschedule22.[Day], dbo.TblVacationschedule22.color, dbo.TblVacationschedule22.DurationID, "
str = str & "                        dbo.TblDurations_details2.Name, dbo.TblDurations_details2.ID, dbo.TblVacationschedule22.[Date], dbo.TblVacationschedule22.DateH,"
str = str & "                        dbo.TblVacationschedule22.ISVac, dbo.TblDurations_details2.FromDate, dbo.TblDurations_details2.ToDate, dbo.TblVacationTypes.NameE AS VACATIONNAMEe,"
str = str & "                        dbo.TblVacationTypes.Name AS VACATIONNAME"
str = str & "  FROM         dbo.Tbldurations2 INNER JOIN"
str = str & "                        dbo.TblDurations_details2 ON dbo.Tbldurations2.ID = dbo.TblDurations_details2.DID INNER JOIN"
str = str & "                        dbo.TblVacationschedule22 ON dbo.Tbldurations2.ID = dbo.TblVacationschedule22.DurationID AND dbo.TblVacationschedule22.[Date] BETWEEN"
str = str & "                        dbo.TblDurations_details2.FromDate AND dbo.TblDurations_details2.ToDate LEFT OUTER JOIN"
str = str & "                        dbo.TblVacationTypes ON dbo.TblVacationschedule22.VacationTypeID = dbo.TblVacationTypes.ID"
str = str & "   Where (1 = 1)"

    FromDate.Visible = True
    ToDate.Visible = True
    FromdateH.Visible = False
    ToDateH.Visible = False
    
    ElseIf Typ = 1 Then
       str = "    SELECT  TblDurations2.id YID , dbo.TblVacationschedule22.Day  ,  TblVacationschedule22.color  ,   dbo.TblVacationschedule22.DurationID, dbo.TblDurations_Details2.Name, dbo.TblDurations_Details2.ID, dbo.TblVacationschedule22.Date   , dbo.TblVacationschedule22.DateH   ,"
       str = str & "   dbo.TblVacationschedule22.isvac , dbo.TblDurations_Details2.FromDate, dbo.TblDurations_Details2.ToDate"
       str = str & "   FROM     dbo.TblDurations2 INNER JOIN  dbo.TblDurations_Details2 ON dbo.TblDurations2.ID = dbo.TblDurations_Details2.DID"
       str = str & "   INNER JOIN  dbo.TblVacationschedule22 ON dbo.TblDurations2.ID = dbo.TblVacationschedule22.DurationID"
       str = str & " and dbo.TblVacationschedule22.DateH >= dbo.TblDurations_Details2.FromDateH And dbo.TblVacationschedule22.DateH <= dbo.TblDurations_Details2.ToDateH"
      ' str = str & " and  format ( dbo.TblVacationschedule22.DateH, 'yyyy/MM/dd') >=  format ( dbo.TblDurations_Details2.FromDateH , 'yyyy/MM/dd') And format ( dbo.TblVacationschedule22.DateH , 'yyyy/MM/dd') <= format ( dbo.TblDurations_Details2.ToDateH , 'yyyy/MM/dd')  "
       str = str & "   Where  1 =1 "
       
    FromDate.Visible = False
    ToDate.Visible = False
    FromdateH.Visible = True
    ToDateH.Visible = True
    
    End If
    
    
    str = str & "   and  TblDurations_Details2.ID = " & DurID
    fg.Rows = 1
    Dim vac As String
    rs.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If rs.RecordCount > 0 Then
    
         rs.MoveFirst
        fg.Rows = rs.RecordCount + 1
        DcYear.BoundText = rs("YID").value
        With fg
 
                For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                         .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                          If rs("isvac").value = True Then
                                vac = "⁄ÿ·…"
                                
                                .Row = i
                                .Col = .ColIndex("vac")
                                .CellBackColor = RGB(0, 255, 0)
                           
                          ElseIf rs("isvac").value = False Then
                                vac = "⁄„· "
                                .Row = i
                                .Col = .ColIndex("vac")
                                .CellBackColor = RGB(0, 0, 0)
                          End If
'                          .TextMatrix(I, .ColIndex("VACATIONNAME")) = IIf(IsNull(rs("VACATIONNAME").value), "", rs("VACATIONNAME").value)
                         .TextMatrix(i, .ColIndex("vac")) = vac
                         .TextMatrix(i, .ColIndex("day")) = IIf(IsNull(rs("Day").value), "", rs("Day").value)       'WeekdayName(Weekday(rs("Date").value, vbSaturday), False, vbSaturday)
                         .TextMatrix(i, .ColIndex("Date")) = Format(IIf(IsNull(rs("Date").value), Date, rs("Date").value), "dd/MM/yyyy")
                         .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), Date, rs("DateH").value)
                          .Cell(flexcpBackColor, i, .ColIndex("vac"), i, .ColIndex("vac")) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                          rs.MoveNext
                Next
        End With
    End If
  

If Typ = 0 Then
        fg.ColWidth(fg.ColIndex("Date")) = 1200
        fg.ColWidth(fg.ColIndex("DateH")) = 1200
ElseIf Typ = 1 Then
        fg.ColWidth(fg.ColIndex("Date")) = 1200
        fg.ColWidth(fg.ColIndex("DateH")) = 1200
End If

End Sub
Public Sub Retrive_Det(DurID As Integer, Typ As Integer)

 Set rs = New ADODB.Recordset
    Dim str As String
    
    If Typ = 0 Then
    str = "   SELECT  tbldurations.id YID , dbo.TblVacationSchedule.Day,  TblVacationSchedule.color , dbo.TblVacationSchedule.DurationID, dbo.TblDurations_Details.Name, dbo.TblDurations_Details.ID, dbo.TblVacationSchedule.Date   , dbo.TblVacationSchedule.DateH   , dbo.TblVacationSchedule.ISVac, "
    str = str & "  dbo.TblDurations_Details.FromDate , dbo.TblDurations_Details.ToDate"
    str = str & "  FROM     dbo.TblDurations INNER JOIN"
    str = str & "  dbo.TblDurations_Details ON dbo.TblDurations.ID = dbo.TblDurations_Details.DID INNER JOIN"
    str = str & "  dbo.TblVacationSchedule ON dbo.TblDurations.ID = dbo.TblVacationSchedule.DurationID"
    str = str & "  and  (dbo.TblVacationSchedule.Date BETWEEN dbo.TblDurations_Details.FromDate AND dbo.TblDurations_Details.ToDate)  "
    str = str & "  Where 1=1   "
    
    FromDate.Visible = True
    ToDate.Visible = True
    FromdateH.Visible = False
    ToDateH.Visible = False
    
    ElseIf Typ = 1 Then
       str = "    SELECT  tbldurations.id YID , dbo.TblVacationSchedule.Day  ,  TblVacationSchedule.color  ,   dbo.TblVacationSchedule.DurationID, dbo.TblDurations_Details.Name, dbo.TblDurations_Details.ID, dbo.TblVacationSchedule.Date   , dbo.TblVacationSchedule.DateH   ,"
       str = str & "   dbo.TblVacationSchedule.isvac , dbo.TblDurations_Details.FromDate, dbo.TblDurations_Details.ToDate"
       str = str & "   FROM     dbo.TblDurations INNER JOIN  dbo.TblDurations_Details ON dbo.TblDurations.ID = dbo.TblDurations_Details.DID"
       str = str & "   INNER JOIN  dbo.TblVacationSchedule ON dbo.TblDurations.ID = dbo.TblVacationSchedule.DurationID"
       str = str & " and dbo.TblVacationSchedule.DateH >= dbo.TblDurations_Details.FromDateH And dbo.TblVacationSchedule.DateH <= dbo.TblDurations_Details.ToDateH"
      ' str = str & " and  format ( dbo.TblVacationSchedule.DateH, 'yyyy/MM/dd') >=  format ( dbo.TblDurations_Details.FromDateH , 'yyyy/MM/dd') And format ( dbo.TblVacationSchedule.DateH , 'yyyy/MM/dd') <= format ( dbo.TblDurations_Details.ToDateH , 'yyyy/MM/dd')  "
       str = str & "   Where  1 =1 "
       
    FromDate.Visible = False
    ToDate.Visible = False
    FromdateH.Visible = True
    ToDateH.Visible = True
    
    End If
    
    
    str = str & "   and  TblDurations_Details.ID = " & DurID
    fg.Rows = 1
    Dim vac As String
    rs.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If rs.RecordCount > 0 Then
    
         rs.MoveFirst
        fg.Rows = rs.RecordCount + 1
        DcYear.BoundText = rs("YID").value
        With fg
 
                For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                         .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                          If rs("isvac").value = True Then
                                vac = "⁄ÿ·…"
                                
                                .Row = i
                                .Col = .ColIndex("vac")
                                .CellBackColor = RGB(0, 255, 0)
                           
                          ElseIf rs("isvac").value = False Then
                                vac = "⁄„· "
                                .Row = i
                                .Col = .ColIndex("vac")
                                .CellBackColor = RGB(0, 0, 0)
                          End If
                         .TextMatrix(i, .ColIndex("vac")) = vac
                         .TextMatrix(i, .ColIndex("day")) = IIf(IsNull(rs("Day").value), "", rs("Day").value)       'WeekdayName(Weekday(rs("Date").value, vbSaturday), False, vbSaturday)
                         .TextMatrix(i, .ColIndex("Date")) = Format(IIf(IsNull(rs("Date").value), Date, rs("Date").value), "dd/MM/yyyy")
                         .TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(rs("DateH").value), Date, rs("DateH").value)
                          .Cell(flexcpBackColor, i, .ColIndex("vac"), i, .ColIndex("vac")) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                          rs.MoveNext
                Next
        End With
    End If
  

If Typ = 0 Then
'        fg.ColWidth(fg.ColIndex("Date")) = 1200
'        fg.ColWidth(fg.ColIndex("DateH")) = 0
ElseIf Typ = 1 Then
'        fg.ColWidth(fg.ColIndex("Date")) = 0
'        fg.ColWidth(fg.ColIndex("DateH")) = 1200
End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
   ' Dim IntResult As String
   ' Dim StrMSG As String
   ' On Error GoTo ErrTrap
''
'    If Me.TxtModFlg.text <> "R" Then
'
'        Select Case Me.TxtModFlg.text
''
''            Case "N"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
''                    StrMSG = "You will close this screen before save " & Chr(13)
'                    StrMSG = StrMSG & " the new data  " & Chr(13)
'                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
''                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
'                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
'
''                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
'                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
''                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
''                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
'
'                End If
'''
 '           Case "E"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save  " & Chr(13)
'                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
'                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
''                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
'                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
''                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
''                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
'                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
'                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
'                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
''
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
'                SaveData
'
                'btnSave
'            Case vbCancel
'                Cancel = True
'        End Select
''
'    End If
'
'    Exit Sub
'ErrTrap:
'
End Sub

Function CuurentLogdata(Optional Currentmode As String)
 
    
End Function

