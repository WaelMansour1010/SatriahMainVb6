VERSION 5.00
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form frmabout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4380
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6465
   ClipControls    =   0   'False
   HelpContextID   =   190
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3023.153
   ScaleMode       =   0  'User
   ScaleWidth      =   6070.969
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      Picture         =   "frmAbout.frx":038A
      RightToLeft     =   -1  'True
      ScaleHeight     =   2055
      ScaleWidth      =   2775
      TabIndex        =   20
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   270
      RightToLeft     =   -1  'True
      ScaleHeight     =   795
      ScaleWidth      =   6135
      TabIndex        =   5
      Top             =   3630
      Width           =   6135
      Begin ImpulseButton.ISButton cmdok 
         Height          =   345
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "„Ê«›ﬁ"
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
      Begin ImpulseButton.ISButton cmdSysInfo 
         Height          =   345
         Left            =   2310
         TabIndex        =   7
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»Ì«‰«  «·‰Ÿ«„"
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
         Height          =   345
         Left            =   4680
         TabIndex        =   8
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         ButtonStyle     =   1
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
      Begin ImpulseAniLabel.ISAniLabel LblLink 
         Height          =   225
         Left            =   3450
         TabIndex        =   9
         Top             =   1620
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   397
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
         MouseIcon       =   "frmAbout.frx":2B6E
         BackColor       =   16777215
         Caption         =   "WWW.BisEgypt.com"
         ColorHover      =   16711680
         ImageCount      =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   5670
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   150
      RightToLeft     =   -1  'True
      ScaleHeight     =   1815
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   1770
      Width           =   6915
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "www.sattaryah.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   -1920
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Êﬁ⁄ «·«·ﬂ —Ê‰Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   4440
         TabIndex        =   16
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "info@sattaryah.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   -1860
         TabIndex        =   15
         Top             =   1110
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(+2)0123591024 - 0114448733 - 0103366467"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   1170
         TabIndex        =   14
         Top             =   3960
         Visible         =   0   'False
         Width           =   49995
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(+966)114030870   (966)114094121"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   900
         TabIndex        =   13
         Top             =   840
         Width           =   3960
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Â« ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·—Ì«÷- «·„„·ﬂ… «·⁄—»Ì… «·”⁄ÊœÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   420
         Width           =   3135
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Õ«”» «·⁄—»Ï «·„ ﬂ«„·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   7080
         TabIndex        =   10
         Top             =   30
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Image ImgCur 
         Height          =   480
         Left            =   750
         Picture         =   "frmAbout.frx":2CD0
         Top             =   2250
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " : «·»—Ìœ «·≈·ﬂ —Ê‰Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   4470
         TabIndex        =   4
         Top             =   1110
         Width           =   1605
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ·Ì›«ﬂ”:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   12
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2850
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " „Ã„Ê⁄Â «·”« —Ì…   - Õ·Ê· „«· Ê «⁄„«·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   2550
         TabIndex        =   2
         Top             =   30
         Width           =   3495
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "36 «„ œ«œ Ê·Ï «·⁄Âœ - Õœ«∆ﬁ «·ﬁ»… - «·ﬁ«Â—… - „’—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Index           =   0
         Left            =   7140
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   4545
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dynamic Byte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   11
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "‰Ÿ«„ œÌ‰«„Ìﬂ »«Ì "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   9
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   105
      Left            =   4440
      Picture         =   "frmAbout.frx":2E22
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   2100
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx _
                Lib "advapi32" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                Lib "advapi32" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          ByRef lpType As Long, _
                                          ByVal lpData As String, _
                                          ByRef lpcbData As Long) As Long

Private Declare Function RegCloseKey _
                Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub CmdOk_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdOk_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim Msg As String
    Resize_Form Me, False

    FormPostion Me, GetPostion
    Me.Caption = "⁄‰ " & App.title

    If Dir(App.path & "\Garphics\About.bmp") <> "" Then
        'PctAbout.Picture = LoadPicture(App.Path & "\Garphics\About.bmp")
    End If

    Me.Image1.Picture = mdifrmmain.ImgLstMenuIcons.ListImages("Web").Picture
    Set Me.CmdHelp.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Help2").Picture
    Set Me.cmdok.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Tick").Picture

    'If Dir(App.Path & "\Garphics\SLogo.jpg") <> "" Then
    '    PicMark.Picture = LoadPicture(App.Path & "\Garphics\SLogo.BMP")
    'End If
    'lblDisclaimer.Caption = Msg

End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
    Dim rc As Long
    Dim SysInfoPath As String
   
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then

        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If

        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, _
                            KeyName As String, _
                            SubKeyRef As String, _
                            ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...

        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value

        Case REG_DWORD                                          ' Double Word Registry Key Data Type

            For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
            Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:              ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub lblDisclaimer_Click()
    OpenWebSite
End Sub

Private Sub lblDisclaimer_MouseEnter()

End Sub

Private Sub lblDisclaimer_MouseLeave()
End Sub

Private Sub LblLink_Click()
    OpenWebSite
End Sub

Private Sub lblTitle_Click(Index As Integer)
    On Error Resume Next

    Select Case Index

        Case 6
            Shell "mailto:" & Trim$(Me.lblTitle(Index).Caption), vbNormalFocus
    End Select

End Sub

Private Sub lblTitle_MouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    Select Case Index

        Case 6
            lblTitle(Index).ForeColor = vbBlue
            lblTitle(Index).FontUnderline = True
            lblTitle(Index).MousePointer = vbCustom
            Set lblTitle(Index).MouseIcon = Me.ImgCur.Picture
    End Select

End Sub

Private Sub PctAbout_Click()

End Sub
