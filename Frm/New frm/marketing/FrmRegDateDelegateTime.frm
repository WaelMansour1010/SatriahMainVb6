VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmRegDateDelegateTime 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈⁄œ«œ«  „Ê«⁄Ìœ «·„‰«œÌ»"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "FrmRegDateDelegateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   10515
   Begin VB.Frame lbldata 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   10545
      Begin VB.ComboBox DcbEnd 
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
      End
      Begin VB.ComboBox DcbTime 
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox DcbStart 
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œﬁÌﬁÂ"
         Height          =   285
         Index           =   2
         Left            =   5400
         TabIndex        =   18
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„· Ì‰ ÂÌ «·”«⁄Â"
         Height          =   285
         Index           =   1
         Left            =   8880
         TabIndex        =   17
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   975
         Left            =   0
         Top             =   270
         Width           =   5055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Â–Â «·‘«‘Â  ﬁÊ„ » ”ÃÌ· «⁄œ«œ«  „Ê«⁄Ìœ «·„‰«œÌ»"
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
         Height          =   780
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„· Ì»œ« «·”«⁄Â"
         Height          =   285
         Index           =   24
         Left            =   9120
         TabIndex        =   15
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÂ «·“„‰ÌÂ"
         Height          =   285
         Index           =   15
         Left            =   9240
         TabIndex        =   14
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13140
      TabIndex        =   12
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10515
      _cx             =   18547
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " ≈⁄œ«œ«  „Ê«⁄Ìœ «·„‰«œÌ»"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   5880
         Picture         =   "FrmRegDateDelegateTime.frx":038A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   9
         Top             =   0
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3345
      _cx             =   5900
      _cy             =   953
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   2
         Top             =   120
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "Œ—ÊÃ"
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13200
      TabIndex        =   3
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   13560
      TabIndex        =   6
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   4
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmRegDateDelegateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String





Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

       ' Case 0

           ' If DoPremis(Do_New, Me.name, True) = False Then
           '     Exit Sub
           ' End If
'
'            TxtModFlg.text = "N"
'            clear_all Me
           ' lbl(20).Caption = "0"
           ' lbl(21).Caption = "0"
           ' lbl(22).Caption = "0"
           ' lbl(23).Caption = "0"
            
              'GRID2.Clear flexClearScrollable, flexClearEverything
   ' GRID2.Rows = 1
   '         Me.DCboUserName.BoundText = user_id
            'TxtPaymentCounts.text = 1
'Dcbranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
'            Accredit.Enabled = True
'                If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
                                               
      '  Case 1
'
'            If DoPremis(Do_Edit, Me.name, True) = False Then
'                Exit Sub
'            End If

'            TxtModFlg.text = "E"
'            Me.DCboUserName.BoundText = user_id

        Case 2
        
'Del_Trans
SaveData
'            Dim Msg As String

'            If Trim(Me.Dcbranch.BoundText) = "" Then
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    Msg = "Specify Branch"
'                Else
'                    Msg = "Õœœ «·›—⁄ "
'                End If

'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                Dcbranch.SetFocus
'                SendKeys "{F4}"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If

'            my_branch = Me.Dcbranch.BoundText

            

'        Case 3
'            Undo

'        Case 4

'            If DoPremis(Do_Delete, Me.name, True) = False Then
'                Exit Sub
'            End If

'            Del_Trans

'        Case 5
'            Load FrmEmpAdvanceSearch
'            FrmEmpAdvanceSearch.Show

        Case 2
     '   Me.e
     '       Unload Me

        'Case 7
        '    ShowGL_cc Me.TxtNoteSerial.text, , 200
'
'        Case 8
           'CalCulateParts
            
            
'                 Case 9

'            If DoPremis(Do_Print, Me.name, True) = False Then
'                Exit Sub
'            End If
'
        '   ' If val(Me.XPTxtID.text) <> 0 Then
        '        print_report val(Me.XPTxtID.text)
        '
        
        '    End If
        
    End Select

    Exit Sub
ErrTrap:
End Sub


Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    Set TTD = New clstooltipdemand
   ' Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
   ' Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
   ' Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
   ' Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    
   ' Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
  '  Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
Dim i, j As Integer
    Set Dcombos = New ClsDataCombos
For i = 0 To 23
If i < 10 Then
Me.DcbEnd.AddItem "0" & i
Me.DcbStart.AddItem "0" & i
Else
Me.DcbEnd.AddItem i
Me.DcbStart.AddItem i
End If
Next i
Me.DcbTime.AddItem "15"
Me.DcbTime.AddItem "30"
Me.DcbTime.AddItem "60"
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblRegTimeDelgate     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '  XPDtbTrans.value = Date
       ' Me.TxtModFlg.text = "R"
    Retrive




    If OPEN_NEW_SCREEN = True Then
        'Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
   ' Set XPic = Me.XPBtnMove(1).ButtonImage
   ' Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
   ' Set XPic = Me.XPBtnMove(0).ButtonImage
   ' Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
   ' Set Me.XPBtnMove(3).ButtonImage = XPic
 '   Label1.Visible = False

   ' Cmd(0).Caption = "New"
   ' Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
  '  Cmd(3).Caption = "Undo"
  '  Cmd(4).Caption = "Delete"
  '  Cmd(5).Caption = "Search"
 'Cmd(9).Caption = "Prinet"
    Cmd(6).Caption = "Exit"
 '   CmdHelp.Caption = "Help"

    Me.Caption = " Settings dates Almnadeb"
    EleHeader.Caption = Me.Caption
    lbl(24).Caption = "Work Start"
    lbl(1).Caption = "Work End"

    lbl(15).Caption = "Period of time"
    lbl(2).Caption = "Minutes"
    
    lbl(25).Caption = "This screen settings you log dates Almnadeb"

    'lbl(0).Caption = "Box"
   ' Fra(0).Caption = "payments Method"
   ' lbl(9).Caption = "Count"
  '  lbl(10).Caption = "Start"
   ' lbl(11).Caption = "Month"
   ' lbl(12).Caption = "Year"
  '  Cmd(8).Caption = "Calc Dates"
  '  ChkSaleryDis.Caption = "Auto Discount"
  '  lbl(8).Caption = "By"
  '  lbl(7).Caption = "Curr rec."
  '  lbl(6).Caption = "rec. count"
'XPTab301.Caption = "Data"
   ' With Me.Fg
   '     .TextMatrix(0, .ColIndex("PartNO")) = "NO"
   '     .TextMatrix(0, .ColIndex("PartValue")) = "Value"
   '     .TextMatrix(0, .ColIndex("PartDate")) = "Date"
'
'    End With

End Sub

'Private Sub YearMonth()

  '  Dim i As Integer
  '  Dim IntDefIndex As Integer

  '  CmbMonth.Clear

    'For i = 1 To 12
     '   CmbMonth.AddItem MonthName(i)
   ' Next

    'CmbMonth.ListIndex = Month(Date) - 1
   ' CboYear.Clear

   ' For i = 2010 To 2050
      '  CboYear.AddItem i

     '   If i = year(Date) Then
      '      IntDefIndex = CboYear.NewIndex
     '   End If

   ' Next

   ' CboYear.ListIndex = IntDefIndex
'End Sub

Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
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

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

'Private Sub TxtAdvanceValue_LostFocus()
'
'
'End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = " ﬁ—Ì— ≈’«»… ⁄„·   "
            Me.Cmd(2).Enabled = True
           ' Me.Cmd(3).Enabled = False
           ' Me.Cmd(0).Enabled = True
           ' Me.Cmd(1).Enabled = True
           ' Me.Cmd(4).Enabled = True
           ' Me.Cmd(5).Enabled = True
         ' '  Me.XPBtnMove(0).Enabled = True
         '   Me.XPBtnMove(1).Enabled = True
         '   Me.XPBtnMove(2).Enabled = True
         '   Me.XPBtnMove(3).Enabled = True
         '   TxtAdvanceValue.Locked = True
            Me.DcboBox.locked = True
         '   XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
              '  Me.XPBtnMove(0).Enabled = False
              '  Me.XPBtnMove(1).Enabled = False
              '  Me.XPBtnMove(2).Enabled = False
              '  Me.XPBtnMove(3).Enabled = False
              '  Me.Cmd(1).Enabled = False
              '  Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '        Me.Caption = " ﬁ—Ì— ≈’«»… ⁄„·   ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
       '     TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            'XPDtbTrans.Enabled = True
            'XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = " ﬁ—Ì— ≈’«»… ⁄„·   (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
           ' Me.XPBtnMove(0).Enabled = False
           ' Me.XPBtnMove(1).Enabled = False
           ' Me.XPBtnMove(2).Enabled = False
           ' Me.XPBtnMove(3).Enabled = False
         '   TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
         '   XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

 


'Private Sub XPBtnMove_Click(Index As Integer)
'    On Error GoTo ErrTrap
'
'    If Me.TxtModFlg.text = "N" Then
'        clear_all Me
 '       Me.TxtModFlg.text = "R"
''        XPBtnMove_Click (1)
 '   End If
''
 '   Select Case Index
'
'        Case 0
'
'            If Not (rs.EOF Or rs.BOF) Then
 '               rs.MovePrevious
'
''                If rs.BOF Then rs.MoveFirst
 '           End If

 '       Case 1

 '           If Not (rs.EOF Or rs.BOF) Then
' '               rs.MoveFirst
'            End If
'
'        Case 2

'            If Not (rs.EOF Or rs.BOF) Then
'                rs.MoveLast
'            End If
'
'        Case 3
'
'            If Not (rs.EOF Or rs.BOF) Then
'                rs.MoveNext
'
'                If rs.EOF Then rs.MoveLast
'            End If
'
'    End Select
'
'    Retrive
'    Exit Sub
'ErrTrap:
'End Sub
'
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
   Dim StrSQL As String

   'On Error GoTo ErrTrap


    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
  Me.DcbStart.text = rs("WorkStart").value
  Me.DcbEnd.text = rs("WorkEnd").value
    Me.DcbTime.text = rs("WorkTime").value
    'XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)

   '
   '
    Set RsDetails = New ADODB.Recordset
   ' StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID=" & val(XPTxtID.text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    '.Clear flexClearScrollable, flexClearEverything
    'Fg.Rows = Fg.FixedRows

    'If Not (RsDetails.BOF Or RsDetails.EOF) Then
      '  RsDetails.MoveFirst
        ' 'g.Rows = Fg.FixedRows + RsDetails.RecordCount

       ' For i = Me.Fg.FixedRows To Fg.Rows - 1
         '   Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
         '   Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
        ''    Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
        '    RsDetails.MoveNext
        'Next i

  '  End If

    RsDetails.Close
    Set RsDetails = Nothing
    
'    fillapprovData
    
    'XPTxtCurrent.Caption = rs.AbsolutePosition
    'XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap
    If val(Me.DcbEnd.text) <= val(Me.DcbStart.text) Then
    MsgBox "ÌÃ» «‰  ﬂÊ‰ ‰Â«Ì… «·⁄„· «ﬂ»— „‰ »œ«Ì Â"
    DcbEnd.SetFocus
Exit Sub
    End If
    
If Me.DcbEnd.text = "" Then
MsgBox "ÌÃ» √Œ Ì«— ‰Â«Ì… «·⁄„·"
DcbEnd.SetFocus
Exit Sub
End If
If Me.DcbStart.text = "" Then
MsgBox "ÌÃ» √Œ Ì«— »œ«Ì… «·⁄„·"
DcbStart.SetFocus
Exit Sub
End If
If Me.DcbTime.text = "" Then
MsgBox "ÌÃ» √Œ Ì«— «·› —… «·“„‰ÌÂ "
DcbTime.SetFocus
Exit Sub
End If

  
 
        Dim RsTest As New ADODB.Recordset
  


        Cn.BeginTrans
        BeginTrans = True

      
StrSQL = "Delete From TblRegTimeDelgate Where ID<> -1"
            Cn.Execute StrSQL, , adExecuteNoRecords
   Dim j As Integer

For i = val(Me.DcbStart.text) To val(Me.DcbEnd.text) - 1
j = 0
'MsgBox i

Do

If j < 60 Then
'MsgBox i & ". " & j
rs.AddNew
rs("WorkEnd").value = val(Me.DcbEnd.text)
rs("WorkStart").value = val(Me.DcbStart.text)
rs("WorkTime").value = val(Me.DcbTime.text)
rs("name").value = i & ". " & j
Else
If i = (val(Me.DcbEnd.text) - 1) Then
rs.AddNew
rs("WorkEnd").value = val(Me.DcbEnd.text)
rs("WorkStart").value = val(Me.DcbStart.text)
rs("WorkTime").value = val(Me.DcbTime.text)
rs("name").value = i + 1 & ". " & 0
End If
End If
j = j + val(Me.DcbTime.text)
rs.update
Loop While j <= 60

Next i

    
    
             
          
        rs.update
       ' Set RsDetails = New ADODB.Recordset
       ' RsDetails.Open "TblEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

      ' For i = Me.Fg.FixedRows To Fg.Rows - 1
       '     RsDetails.AddNew
       ''     RsDetails("AdvanceID").value = val(XPTxtID.text)
          '  RsDetails("PartNO").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
         '   RsDetails("PartValue").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
          ''  RsDetails("PartDate").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
          '  RsDetails.update
       ' Next i
    
'        Dim NoteID As Long
'        Dim line_no As Integer
'        Dim RsNotes As New ADODB.Recordset
'        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
'        If detect_employee_work_type = 1 Then
        
'            If Me.TxtModFlg.text = "E" Then
 
'                StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords

'            End If

'            RsNotes.AddNew
'            NoteID = CStr(TxtNoteID.text)
'            RsNotes("NoteID").value = CStr(TxtNoteID.text)
'            RsNotes("NoteType").value = 8032
'            RsNotes("NoteDate").value = XPDtbTrans.value
'            RsNotes("UserID").value = user_id
'            RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
'            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            RsNotes("numbering_type1").value = sand_numbering_type(32) ' ”ÃÌ· «·”·›'‰Ê⁄  —ﬁÌ„    
'            RsNotes("sanad_year").value = year(XPDtbTrans.value)
'            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'            RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
            '     RsNotes("remark").value = txtRemarks.text & bankDes
'            RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
                
'            RsNotes.update
                
'            line_no = 1
        
'            Msg = "”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
            '        StrAccountCode = "a1a3a4" 'Õ”«» “„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.TxtAdvanceValue.text), 0, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If

'            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.TxtAdvanceValue.text), 1, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
        
'        End If
    
        Cn.CommitTrans
        BeginTrans = False
'        RsDetails.Close
        Set RsDetails = Nothing
          MsgBox "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
          Unload Me
        'XPTxtCurrent.Caption = rs.AbsolutePosition
        'XPTxtCount.Caption = rs.RecordCount
    '
    '    Select Case Me.TxtModFlg.text

    '        Case "N"
              
'              Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

    '            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
    '                Cmd_Click (0)
    '                Exit Sub
    '            End If
'
'            Case "E"
'                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        End Select

'        TxtModFlg.text = "R"
'    End If

'    Exit Sub
'ErrTrap:

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


End Sub
'
'Private Sub Undo()
'    On Error GoTo ErrTrap
'
'    Select Case TxtModFlg.text
'
'        Case "N"
'            clear_all Me
'            Me.TxtModFlg.text = "R"
'            XPBtnMove_Click (1)
'
 '       Case "E"
''           ' rs.Find "ID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

'            If rs.EOF Or rs.BOF Then
'                Me.TxtModFlg.text = "R"
'                Exit Sub
'            End If
'
'            Retrive
'            Me.TxtModFlg.text = "R"
'    End Select
'
'    Exit Sub
'ErrTrap:
'End Sub
'
'Private Sub Del_Trans()
'    Dim Msg As String
'    Dim StrSQL As String
'
'    On Error GoTo ErrTrap
'
'    If XPTxtID.text <> "" Then
'        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
'        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
'
'        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
'            If Not rs.RecordCount < 1 Then
'                rs.Delete
'                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
'                rs.MoveFirst
'
'                If rs.RecordCount < 1 Then
'                    clear_all Me
'                    TxtModFlg_Change
'                    XPTxtCurrent.Caption = 0
'                    XPTxtCount.Caption = 0
'                Else
'                    Retrive
'                End If
'            End If
'        End If
'
'    Else
'        clear_all Me
'        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        TxtModFlg_Change
'        Exit Sub
'    End If
'
'    TxtModFlg_Change
'    Exit Sub
'ErrTrap:
'    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
'    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
'    rs.CancelUpdate
'End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.name
                RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
              '    RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
             '      RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                rs1.MoveNext
            Next i

    End If
    
    

End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
'
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'   Else
'    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'    End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
'          Else
'             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
'          End If
'            If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            Else
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            End If
'            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
'          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
'
'
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then
'
'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.BackColor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.BackColor = &HFFFFC0
'        End If
'
'End If
'
'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close

'End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
         '   XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
         '   XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
         '   XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
         '   XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub



'Private Sub Form_QueryUnload(Cancel As Integer, _
'                             UnloadMode As Integer)
'    Dim IntResult As String
'    Dim StrMSG As String
'    On Error GoTo ErrTrap
'
''    If Me.TxtModFlg.text <> "R" Then
'
'        Select Case Me.TxtModFlg.text

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
'            Case "E"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save  " & Chr(13)
'                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
'                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
'                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
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
'                SaveData
'
                ' btnSave
'            Case vbCancel
'                Cancel = True
'        End Select
'
'    End If
'
'    Exit Sub
'ErrTrap:
'End Sub

'Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
 
'End Sub

'Private Function CheckDate() As Boolean
 

 
'End Function

 

