VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAsset 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «·Œ“‰ Ê«·»‰Êﬂ Ê«·⁄Âœ"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "FrmAsset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox XPTxtBoxID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   750
      Width           =   2865
   End
   Begin VB.TextBox XPTxtBoxName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   2865
   End
   Begin VB.TextBox XPMTxtRemark 
      Alignment       =   1  'Right Justify
      Height          =   1035
      Left            =   1380
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1500
      Width           =   2865
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   5355
      _cx             =   9446
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
      Caption         =   "«·Œ“‰ Ê«·»‰Êﬂ"
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   5
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmAsset.frx":038A
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
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmAsset.frx":0724
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
         Height          =   345
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmAsset.frx":0ABE
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
         Height          =   345
         Index           =   3
         Left            =   615
         TabIndex        =   8
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmAsset.frx":0E58
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
      Left            =   4590
      TabIndex        =   9
      Top             =   3210
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   3840
      TabIndex        =   10
      Top             =   3210
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   3105
      TabIndex        =   11
      Top             =   3210
      Width           =   705
      _ExtentX        =   1244
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
      Left            =   2355
      TabIndex        =   12
      Top             =   3210
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   1575
      TabIndex        =   13
      Top             =   3210
      Width           =   765
      _ExtentX        =   1349
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
      Index           =   6
      Left            =   0
      TabIndex        =   14
      Top             =   3210
      Width           =   705
      _ExtentX        =   1244
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   375
      Left            =   750
      TabIndex        =   15
      Top             =   3210
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ "
      Height          =   315
      Index           =   3
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   1
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1500
      Width           =   975
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ "
      Height          =   285
      Index           =   0
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2760
      Width           =   825
   End
End
Attribute VB_Name = "FrmAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Cmd_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "N"
        clear_all Me
        XPTxtBoxID.text = CStr(new_id("tblBoxesData", "BoxID", "", True))
        XPTxtBoxName.SetFocus
    Case 1
        If DoPremis(Do_Edit, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "E"
    Case 2
        SaveData
    Case 3
        Undo
    Case 4
        If DoPremis(Do_Delete, Me.name, True) = False Then
            Exit Sub
        End If
        Del_Company
    Case 5
    Case 6
        Unload Me
End Select
Exit Sub
ErrTrap:
End Sub
Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub
Private Sub Form_Activate()
XPTxtBoxID.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
        XPBtnMove_Click (2)
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        XPBtnMove_Click (1)
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        XPBtnMove_Click (3)
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        XPBtnMove_Click (0)
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
Private Sub Form_Load()
On Error GoTo ErrTrap
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If
Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
Resize_Form Me
AddTip
Set Rs = New ADODB.Recordset
Rs.Open "tblBoxesData", Cn, adOpenStatic, adLockOptimistic, adCmdTable
Me.TxtModFlg.text = "R"
XPBtnMove_Click 2
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IntResult As String
Dim StrMSG  As String
On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
Select Case Me.TxtModFlg.text
    Case "N"
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    Case "E"
        StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
        StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
        StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
        StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
        StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
        StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
End Select
IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
Select Case IntResult
    Case vbYes
        Cancel = True
        SaveData
    Case vbCancel
        Cancel = True
End Select
End If
Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
Dim XPic As IPictureDisp

Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic

Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic

Me.Caption = "Boxes Data"
EleHeader.Caption = Me.Caption
Lbl(0).Caption = "Box Code"
Lbl(3).Caption = "Box Name"
Lbl(1).Caption = "Remarks"
Lbl(2).Caption = "Current Record"
Lbl(4).Caption = "NO. Recordes"

Me.Cmd(0).Caption = "New"
Me.Cmd(1).Caption = "Edit"
Me.Cmd(2).Caption = "Save"
Me.Cmd(3).Caption = "Undo"
Me.Cmd(4).Caption = "Delete"
'Me.Cmd(5).Caption = "Search"
Me.Cmd(6).Caption = "Exit"
'Me.Cmd(7).Caption = "Print"
Me.CmdHelp.Caption = "Help"

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
If Rs.State = adStateOpen Then
    If Not (Rs.EOF Or Rs.BOF) Then
        If Rs.EditMode <> adEditNone Then
            Rs.CancelUpdate
        End If
    End If
    Rs.Close
End If
Set Rs = Nothing
Set TTP = Nothing
Exit Sub
ErrTrap:
End Sub

Private Sub ISBÃœÌœ_Click(Index As Integer)

End Sub

Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "»Ì«‰«  «·Œ“‰ Ê«·»‰Êﬂ Ê«·⁄Âœ"
        Else
            Me.Caption = "Boxes Data"
        End If
        Me.Cmd(2).Enabled = False
        Me.Cmd(3).Enabled = False
        
        Me.Cmd(0).Enabled = True
        Me.Cmd(1).Enabled = True
        Me.Cmd(4).Enabled = True
        
        Me.XPBtnMove(0).Enabled = True
        Me.XPBtnMove(1).Enabled = True
        Me.XPBtnMove(2).Enabled = True
        Me.XPBtnMove(3).Enabled = True
        
        Me.XPTxtBoxID.Locked = True
        Me.XPTxtBoxName.Locked = True
        Me.XPMTxtRemark.Locked = True
        If Rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        End If
    Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "»Ì«‰«  «·Œ“‰ Ê«·»‰Êﬂ Ê«·⁄Âœ( ÃœÌœ )"
        Else
            Me.Caption = "Boxes Data(New)"
        End If
        
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "»Ì«‰«  «·Œ“‰ Ê«·»‰Êﬂ Ê«·⁄Âœ( ÃœÌœ )"
        Else
            Me.Caption = "Boxes Data(New)"
        End If
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        Me.XPTxtBoxID.Locked = True
        Me.XPTxtBoxName.Locked = False
        Me.XPMTxtRemark.Locked = False
    Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.Caption = "»Ì«‰«  «·Œ“‰ Ê«·»‰Êﬂ Ê«·⁄Âœ(  ⁄œÌ· )"
        Else
            Me.Caption = "Boxes Data(Edit)"
        End If
        
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        Me.XPTxtBoxID.Locked = True
        Me.XPTxtBoxName.Locked = False
        Me.XPMTxtRemark.Locked = False
End Select
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
XPTxtBoxID.text = IIf(IsNull(Rs("BoxID").Value), "", Val(Rs("BoxID").Value))
XPTxtBoxName.text = IIf(IsNull(Rs("BoxName").Value), "", Trim(Rs("BoxName").Value))
XPMTxtRemark.text = IIf(IsNull(Rs("Comments").Value), "", Trim(Rs("Comments").Value))
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
Private Sub SaveData()
Dim Msg As String
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset
Dim RsTempM As New ADODB.Recordset
Dim BeginTrans As Boolean
On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
    If XPTxtBoxName.text = "" Then
        MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·Œ“‰… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtBoxName.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg.text
        Case "N"
            StrSQL = "select * From  tblBoxesData where BoxName='" & Trim(XPTxtBoxName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
                Msg = "Â‰«ﬂ Œ“‰… „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBoxName.SetFocus
                Exit Sub
            End If
        Case "E"
            StrSQL = "select * From  tblBoxesData where BoxName='" & Trim(XPTxtBoxName.text) & "'"
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If RsTemp.RecordCount > 0 Then
            If RsTemp("BoxID").Value <> Val(XPTxtBoxID.text) Then
                Msg = "Â‰«ﬂ Œ“‰…  „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
                Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & Chr(13)
                Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·Œ“‰…"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtBoxName.SetFocus
                Exit Sub
            End If
            End If
    End Select
    Select Case Me.TxtModFlg.text
    Case "N"
        Rs.AddNew
    End Select
    Cn.BeginTrans
    BeginTrans = True
    Rs("BoxID").Value = Val(XPTxtBoxID.text)
    Rs("BoxName").Value = Trim(XPTxtBoxName.text)
    Rs("Comments").Value = IIf(XPMTxtRemark.text = "", Null, Trim(XPMTxtRemark.text))
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        If Me.TxtModFlg.text = "N" Then
            Rs("Account_Code").Value = ModAccounts.AddNewAccount("a1a2a1", Trim$(Me.XPTxtBoxName.text), True, False)
        Else
            If Not IsNull(Rs("Account_Code").Value) Then
                ModAccounts.EditAccount Rs("Account_Code").Value, Me.XPTxtBoxName.text
            End If
        End If
    End If
    Rs.update
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = Rs.AbsolutePosition
    XPTxtCount.Caption = Rs.RecordCount
    Select Case Me.TxtModFlg.text
        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·Œ“‰…" & Chr(13)
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
Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         clear_all Me
         Me.TxtModFlg.text = "R"
         XPBtnMove_Click (1)
    Case "E"
         Rs.Find "BoxID='" & Val(XPTxtBoxID.text) & "'", , adSearchForward, adBookmarkFirst
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
Private Sub Del_Company()
Dim Msg As String
Dim StrSQL As String
Dim RsTemp As New ADODB.Recordset
On Error GoTo ErrTrap

If XPTxtBoxID.text <> "" Then
    StrSQL = "select * From Notes where BoxID=" & Trim(XPTxtBoxID.text)
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Chr(13)
        Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·Œ“‰…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    Msg = "”Ì „ Õ–› »Ì«‰«  «·Œ“‰… —ﬁ„ " & Chr(13)
    Msg = Msg + (XPTxtBoxID.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
        If Not Rs.RecordCount < 1 Then
            Dim StrAccountCode As String
            StrAccountCode = Rs("Account_Code").Value
            If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                Rs.Delete
            Else
                Exit Sub
            End If
            Rs.MoveFirst
            If Rs.RecordCount < 1 Then
                clear_all Me
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
            Else
                Retrive
            End If
        End If
    End If
Else
    clear_all Me
    Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtModFlg_Change
    Exit Sub
End If
TxtModFlg_Change
Exit Sub
ErrTrap:
If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & Chr(13) & Err.Description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
    Rs.CancelUpdate
End If
End Sub
Private Sub AddTip()
Dim Wrap As String
On Error GoTo ErrTrap
Set TTP = New clstooltip
Wrap = Chr(13) + Chr(10)

With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(0), _
    "ÃœÌœ ..." & Wrap & _
    "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(1), _
    " ⁄œÌ· ..." & Wrap & _
    "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(2), _
    "Õ›Ÿ ..." & Wrap & _
    "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & _
     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
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
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(4), _
    "Õ–› ..." & Wrap & _
    "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(5), _
    "»ÕÀ ..." & Wrap & _
    "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & _
    "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(6), _
    "Œ—ÊÃ ..." & Wrap & _
    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(1), _
    "«·√Ê· ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(0), _
    "«·”«»ﬁ ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(3), _
    "«· «·Ì ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(2), _
    "«·√ŒÌ— ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hwnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl CmdHelp, _
    "„”«⁄œ… ..." & Wrap & _
    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
    "≈÷€ÿ Â‰«" & Wrap, True
End With
Exit Sub
ErrTrap:
End Sub





