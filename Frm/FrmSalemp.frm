VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSalemp 
   Caption         =   "Form2"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   ScaleHeight     =   5040
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbMonth 
      Height          =   315
      Left            =   3840
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox CboYear 
      Height          =   315
      Left            =   2280
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox XpTexSalry 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   3960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2280
      Width           =   885
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox XPTxtSalr 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   780
      Width           =   885
   End
   Begin VB.TextBox XPTxtSAlNump 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3960
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1275
      Width           =   885
   End
   Begin VB.TextBox XPMTxtNots 
      Alignment       =   1  'Right Justify
      Height          =   435
      Left            =   1680
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3690
      Width           =   3165
   End
   Begin VB.TextBox XPTxtID 
      Height          =   285
      Left            =   810
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2340
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial 
      Height          =   285
      Left            =   1050
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2940
      Width           =   375
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5835
      _cx             =   10292
      _cy             =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      Caption         =   "œ›⁄ —« » „ÊŸ› "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1155
         TabIndex        =   1
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSalemp.frx":0000
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   2
         Left            =   90
         TabIndex        =   2
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSalemp.frx":039A
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSalemp.frx":0734
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   3
         Left            =   615
         TabIndex        =   4
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmSalemp.frx":0ACE
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   4950
      TabIndex        =   5
      Top             =   4395
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   4395
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
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
      Index           =   2
      Left            =   3195
      TabIndex        =   7
      Top             =   4395
      Width           =   795
      _ExtentX        =   1402
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
      Index           =   3
      Left            =   2325
      TabIndex        =   8
      Top             =   4395
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   4380
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      Index           =   5
      Left            =   570
      TabIndex        =   10
      Top             =   4395
      Width           =   795
      _ExtentX        =   1402
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
   Begin MSDataListLib.DataCombo DcboBoxSal 
      Height          =   315
      Left            =   1890
      TabIndex        =   17
      Top             =   3180
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DBCompEmpl 
      Height          =   315
      Left            =   1920
      TabIndex        =   25
      Top             =   1740
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbsalry 
      Height          =   315
      Left            =   1200
      TabIndex        =   30
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57999361
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   8
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   735
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”‰…"
      Height          =   315
      Index           =   7
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2760
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—« » ‘Â— "
      Height          =   315
      Index           =   6
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—« »"
      Height          =   315
      Index           =   5
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2280
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂÊœ"
      Height          =   315
      Index           =   2
      Left            =   4890
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   780
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1275
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   0
      Left            =   4890
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3690
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊŸ›"
      Height          =   315
      Index           =   3
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label LngDevID 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1290
      TabIndex        =   20
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   375
      Left            =   2730
      TabIndex        =   19
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   255
      Index           =   4
      Left            =   4890
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3210
      Width           =   915
   End
End
Attribute VB_Name = "FrmSalemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Cmd_Click(Index As Integer)
Select Case Index
    Case 0
        TxtModFlg.text = "N"
        clear_all Me
        XPTxtSalr.text = CStr(new_id("TblEmpSalary", "SalaryCode", "", True))
        XPTxtSalr.SetFocus
    Case 1
        TxtModFlg.text = "E"
    Case 2
        SaveData
    Case 3
       Call Undo
    Case 4
'       Del_AssetType
    Case 5
    Case 6
        Unload Me
End Select
Exit Sub
ErrTrap:
End Sub
Public Sub clear_all(Frm As Form)
Dim Ctl As Control
On Error Resume Next
For Each Ctl In Frm.Controls
Debug.Print Ctl.name
    If TypeOf Ctl Is ComboBox Then If Ctl.Tag <> "not" Then Ctl.ListIndex = -1
    If TypeOf Ctl Is OptionButton Then If Ctl.Tag <> "not" Then Ctl.Value = False
    If TypeOf Ctl Is CheckBox Then If Ctl.Tag <> "not" Then Ctl.Value = False
    If TypeOf Ctl Is DataCombo Then If Ctl.Tag <> "not" Then Ctl.BoundText = ""
    If TypeOf Ctl Is TextBox And Ctl.name <> "TxtModFlg" Then Ctl.text = ""
    If TypeOf Ctl Is DTPicker Then Ctl.Value = Date
'    If TypeOf Ctl Is XPDatePicker30 Then Ctl.CurrentDate = ""
    If TypeOf Ctl Is vsFlexGrid Then
        If Ctl.Rows > 1 Then
            Ctl.Clear 1, 1
            Ctl.FixedRows = 1
            Ctl.Rows = Ctl.FixedRows + 1
        End If
    End If
Next
End Sub

Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
Dim StrSQL As String
Dim GrdBack As ClsBackGroundPic

On Error GoTo ErrTrap
Set GrdBack = New ClsBackGroundPic

'With Me.TxtModFlg
'    .RowHeightMin = 300
'    .WallPaper = GrdBack.Picture
'    .AutoSize 0, .Cols - 1, False
'End With

Set TTD = New clstooltipdemand
Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
'Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
Resize_Form Me
AddTip
Set Dcombos = New ClsDataCombos
Dcombos.GetBoxes Me.DcboBoxSal
'Dcombos.GetUsers Me.DCboUserName
Dcombos.GetEmployees Me.DBCompEmpl
SetDtpickerDate Me.XPDtbsalry
YearMonth
Set Rs = New ADODB.Recordset
StrSQL = "select * From TblEmpAdvance  Where (TblEmpAdvance.AdvanceType =0) Order By AdvanceID"
Rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
XPDtbTrans.Value = Date
'Retrive
Me.TxtModFlg.text = "R"
Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional LngID As Long = 0)
On Error GoTo ErrTrap
If Rs.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
XPTxtSalr.text = IIf(IsNull(Rs("SalaryCode").Value), "", Val(Rs("SalaryCode").Value))
XPTxtSAlNump.text = IIf(IsNull(Rs("PYSalery").Value), "", Trim(Rs("PYSalery").Value))
XPMTxtNots.text = IIf(IsNull(Rs("Nots").Value), "", Trim(Rs("Nots").Value))
XPTxtCurrent.Caption = Rs.AbsolutePosition
XPTxtCount.Caption = Rs.RecordCount
Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MovePrevious
            If Rs.BOF Then Rs.MoveFirst
        End If
    Case 1
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveFirst
        End If
    Case 2
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveLast
        End If
    Case 3
        If Not (Rs.EOF Or Rs.BOF) Then
            Rs.MoveNext
            If Rs.EOF Then Rs.MoveLast
        End If
End Select
Retrive
Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
        Me.Caption = "«·«’Ê· «·À«» …"
        Me.Cmd(2).Enabled = False
        Me.Cmd(3).Enabled = False
        
        Me.Cmd(0).Enabled = True
        Me.Cmd(1).Enabled = True
        Me.Cmd(4).Enabled = True
        
        Me.XPBtnMove(0).Enabled = True
        Me.XPBtnMove(1).Enabled = True
        Me.XPBtnMove(2).Enabled = True
        Me.XPBtnMove(3).Enabled = True
        
        Me.XPTxtSalr.Locked = True
'        Me.XPTxtBankName.Locked = True
        Me.XPMTxtNots.Locked = True
        If Rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        End If
    Case "N"
        Me.Caption = "√‰Ê«⁄ «·„’—Ê›« ( ÃœÌœ )"
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        Me.XPTxtSalr.Locked = True
'        Me.XPTxtBankName.Locked = False
        Me.XPMTxtNots.Locked = False
    Case "E"
        Me.Caption = "√‰Ê«⁄ «·„’—Ê›« (  ⁄œÌ· )"
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        Me.XPTxtSalr.Locked = True
'        Me.XPTxtBankName.Locked = False
        Me.XPMTxtNots.Locked = False
End Select
Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         clear_all Me
         Me.TxtModFlg.text = "R"
         XPBtnMove_Click (1)
    Case "E"
         Rs.Find "SalaryCode=" & Val(XPTxtSalr.text) & "", , adSearchForward, adBookmarkFirst
         If Rs.EOF Or Rs.BOF Then
            Me.TxtModFlg.text = "R"
            Exit Sub
         End If
         Retrive
         Me.TxtModFlg.text = "R"
End Select
Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
Dim Msg As String
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset
Dim RsTempM As New ADODB.Recordset
Dim BeginTrans As Boolean
'On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
    If XPTxtBankName.text = "" Then
        MsgBox "„‰ ›÷·ﬂ √œŒ· ‰Ê⁄ «·«’· ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtBankName.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg.text
        Case "N"
                XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))

            StrSQL = "select * From  AssetType where Name='" & Trim(XPTxtBankName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
                Msg = "Â‰«ﬂ ‰Ê⁄ «’Ê· „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBankName.SetFocus
                Exit Sub
            End If
        Case "E"
        
            StrSQL = "select * From  AssetType where Name='" & Trim(XPTxtBankName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
            If RsTemp("ID").Value <> Val(XPTxtBankID.text) Then
                Msg = "Â‰«ﬂ ‰Ê⁄ „’—Ê›«  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ ‰Ê⁄ «·„’—Ê›«  «·„Õœœ"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBankName.SetFocus
                Exit Sub
            End If
            End If
    End Select
    Cn.BeginTrans
    BeginTrans = True
    Select Case Me.TxtModFlg.text
        Case "N"
                XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
                Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=3"))
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Rs.AddNew
            Rs("ID").Value = Val(XPTxtBankID.text)
        Case "E"
          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End Select
    Rs("Name").Value = Trim(XPTxtBankName.text)
    Rs("Remarks").Value = IIf(XPMTxtRemark.text = "", Null, Trim(XPMTxtRemark.text))
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        If Me.TxtModFlg.text = "N" Then
            Rs("Account_Code").Value = ModAccounts.AddNewAccount("a1a1", Trim$(Me.XPTxtBankName.text), True, False)
        Else
            If Not IsNull(Rs("Account_Code").Value) Then
                ModAccounts.EditAccount Rs("Account_Code").Value, Me.XPTxtBankName.text
            End If
        End If
    End If
    Rs.update
    Cn.CommitTrans
    '*********************************************
           Dim RsNotes As New ADODB.Recordset
        Set RsNotes = New ADODB.Recordset
        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    RsNotes.AddNew
    RsNotes("NoteID").Value = Val(XPTxtID.text)
    RsNotes("NoteSerial").Value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    RsNotes("Note_Value").Value = IIf(XPTxtVal.text = "", Null, Val(XPTxtVal.text))
    'RsNotes("Remark").Value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
    RsNotes("BankID").Value = Null
    RsNotes("CusID").Value = Null
    RsNotes("NoteType").Value = 3
    RsNotes("NoteDate").Value = Date
    RsNotes("UserID").Value = User_ID
    RsNotes("ExpensesID").Value = XPTxtBankID 'IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
    RsNotes("BoxID").Value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
  '  RsNotes("bakhID").Value = IIf(CboBakhara.BoundText = "", Null, Val(CboBakhara.BoundText))
    RsNotes.update

    
   '************************«·ﬁÌÊœ*****************
       LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
       Dim RsDev As New ADODB.Recordset
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '«·ÿ—› «·„œÌ‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").Value = LngDevID
            RsDev("DEV_ID_Line_No").Value = 1
            RsDev("Account_Code").Value = Rs("Account_Code").Value 'Me.DcboDebitSide.BoundText
            RsDev("Value").Value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").Value = 0
'            RsDev("Double_Entry_Vouchers_Description").Value = XPMTxtRemarks.text
            RsDev("Notes_ID").Value = Val(XPTxtID.text)
            RsDev("RecordDate").Value = Date ' Me.XPDtbTrans.Value
'            RsDev("UserID").Value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").Value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update
        '«·ÿ—› «·œ«∆‰
        RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").Value = LngDevID
            RsDev("DEV_ID_Line_No").Value = 2
            RsDev("Account_Code").Value = DcboBox.BoundText ' Me.DcboCreditSide.BoundText
            RsDev("Value").Value = Val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").Value = 1
'            RsDev("Double_Entry_Vouchers_Description").Value = XPMTxtRemarks.text
            RsDev("RecordDate").Value = Date 'Me.XPDtbTrans.Value
            RsDev("Notes_ID").Value = Val(XPTxtID.text)
'            RsDev("UserID").Value = Me.DCboUserName.BoundText
'            RsDev("Account_Interval_ID").Value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update
      '  LblDevID.Caption = LngDevID
'**************************************************************************
    
    
    
    
    
    BeginTrans = False
    XPTxtCurrent.Caption = Rs.AbsolutePosition
    XPTxtCount.Caption = Rs.RecordCount
    Select Case Me.TxtModFlg.text
        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·‰Ê⁄" & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            Cmd_Click (0)
            Exit Sub
            End If
            
        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End Select
    TxtModFlg.text = "R"
End If
Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Rs.EditMode <> adEditNone Then
        Rs.CancelUpdate
    End If
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub AddTip()
Dim Wrap As String
On Error GoTo ErrTrap
Set TTP = New clstooltip
Wrap = Chr(13) + Chr(10)
If SystemOptions.UserInterface = ArabicInterface Then
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(0), _
            "ÃœÌœ ..." & Wrap & _
            "·«÷«›… »Ì«‰«  ‰Ê⁄ ÃœÌœ" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(1), _
            " ⁄œÌ· ..." & Wrap & _
            "· ⁄œÌ· »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(2), _
            "Õ›Ÿ ..." & Wrap & _
            "·Õ›Ÿ »Ì«‰«  «·‰Ê⁄ «·ÃœÌœ" & Wrap & _
             "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(3), _
            " —«Ã⁄ ..." & Wrap & _
            "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
             "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
         With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(4), _
            "Õ–› ..." & Wrap & _
            "·Õ–› »Ì«‰«  Â–« «·‰Ê⁄" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl Cmd(6), _
            "Œ—ÊÃ ..." & Wrap & _
            "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(1), _
            "«·√Ê· ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(0), _
            "«·”«»ﬁ ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(3), _
            "«· «·Ì ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
        With TTP
           .Create Me.hwnd, "√‰Ê«⁄ «·„’—Ê›« ", 1, 15204351, -2147483630
           .MaxWidth = 4000
           .VisibleTime = 9000
           .DelayTime = 600
           .AddControl XPBtnMove(2), _
            "«·√ŒÌ— ..." & Wrap & _
            "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
            " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With
ElseIf SystemOptions.UserInterface = EnglishInterface Then

End If
Exit Sub
ErrTrap:
End Sub





Private Sub YearMonth()

Dim I As Integer
Dim IntDefIndex As Integer

CmbMonth.Clear
For I = 1 To 12
    CmbMonth.AddItem MonthName(I)
Next
CmbMonth.ListIndex = Month(Date) - 1
CboYear.Clear
For I = 2000 To 2050
    CboYear.AddItem I
    If I = Year(Date) Then
        IntDefIndex = CboYear.NewIndex
    End If
Next
CboYear.ListIndex = IntDefIndex
End Sub

