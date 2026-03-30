VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmDailyToolTip 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ·„ÌÕ «·ÌÊ„"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "FrmDailyToolTip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   6615
   Begin XtremeSuiteControls.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6615
      _Version        =   786432
      _ExtentX        =   11668
      _ExtentY        =   8070
      _StockProps     =   173
      BackColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin ImpulseButton.ISButton ISButton2 
      Height          =   390
      Left            =   4320
      TabIndex        =   4
      Top             =   5100
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   688
      ButtonStyle     =   1
      Caption         =   "«· «·Ì"
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
      ButtonImage     =   "FrmDailyToolTip.frx":038A
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
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   390
      Left            =   5460
      TabIndex        =   3
      Top             =   5100
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«·”«»ﬁ"
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
      ButtonImage     =   "FrmDailyToolTip.frx":0724
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
   Begin VB.CheckBox ChkShow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«  ⁄—÷Â« „—… √Œ—Ï"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4620
      Width           =   1785
   End
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   5100
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   688
      ButtonStyle     =   1
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
      BackStyle       =   0
      ButtonImage     =   "FrmDailyToolTip.frx":0ABE
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
   Begin ImpulseAniLabel.ISAniLabel LblLink 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   4620
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   503
      ActiveUnderline =   -1  'True
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   4210688
      MousePointer    =   99
      MouseIcon       =   "FrmDailyToolTip.frx":0E58
      BackColor       =   14871017
      Alignment       =   1
      Caption         =   " «·„“Ìœ „‰ «·„⁄·Ê„«  "
      ColorHover      =   16711680
      ImageCount      =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6585
      X2              =   0
      Y1              =   4980
      Y2              =   4995
   End
End
Attribute VB_Name = "FrmDailyToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrText As String
    Dim StrFileName As String
    Dim StrSQL As String
    Dim BolShowToolTip As Boolean
    CenterForm Me
    Set rs = New ADODB.Recordset
    'StrSql = "select * From ToolTip where HadShow= False"
    rs.Open "ToolTip", Cn, adOpenStatic, adLockOptimistic, adCmdTable
HadnotShow:

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        rs.find "HadShow=False", , adSearchForward, adBookmarkFirst
    Else
        rs.find "HadShow=0", , adSearchForward, adBookmarkFirst
    End If

    If rs.EOF Or rs.BOF Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "update  ToolTip set HadShow=False"
        Else
            StrSQL = "update  ToolTip set HadShow=0"
        End If

        Cn.Execute StrSQL
        rs.Requery
        GoTo HadnotShow
    End If

    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = New ADODB.Recordset
        rs.Open "ToolTip", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        rs.MoveFirst
    End If

    ShowData
    BolShowToolTip = GetSetting(StrAppRegPath, "View_Type", "ShowToolTip", True)

    If BolShowToolTip = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "ShowToolTip", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "ShowToolTip", True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ISButton1_Click()
    On Error GoTo ErrTrap

    If Not (rs.EOF Or rs.BOF) Then
        rs.MovePrevious

        If rs.BOF Then rs.MoveFirst
        ShowData
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ISButton2_Click()
    On Error GoTo ErrTrap

    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveNext

        If rs.EOF Then rs.MoveLast
        ShowData
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub LblLink_Click()
    SystemOptions.SysHelp.HHTopicID = val(LblLink.Tag)
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub ShowData()
    On Error GoTo ErrTrap
    Dim StrFileName As String
    LblLink.Tag = IIf(IsNull(rs("HelpID").value), "", rs("HelpID").value)

    If LblLink.Tag = "" Then
        LblLink.Enabled = False
    Else
        LblLink.Enabled = True
    End If

    StrFileName = App.path & "\DailyToolTip\Temp.html"

    If Dir(StrFileName, vbNormal) <> "" Then
        Kill StrFileName
    End If

    Open StrFileName For Output As #1
    Print #1, rs("ToolTipText").value
    Debug.Print rs("ToolTipID").value
    Close #1
    rs("HadShow").value = True
    rs.update
    WebBrowser1.Navigate StrFileName
    Exit Sub
ErrTrap:
End Sub
