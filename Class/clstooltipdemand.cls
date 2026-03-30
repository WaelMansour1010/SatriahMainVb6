VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstooltipdemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls _
                Lib "comctl32.dll" ()

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private HwndCol() As Long
Const ws = &H400000
Const gw = -20

''Windows API Functions
Private Declare Function CreateWindowEx _
                Lib "user32.dll" _
                Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal x As Long, _
                                         ByVal Y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByRef lParam As Any) As Long

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Declare Function DestroyWindow _
                Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ClientToScreen _
                Lib "user32" (ByVal hWnd As Long, _
                              lpPoint As POINTAPI) As Long

''Windows API Constants
Private Const WM_USER = &H400

Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

'Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2

Private Const TTF_TRANSPARENT = &H100

Private Const TTF_CENTERTIP = &H2

Private Const TTM_ADDTOOLA = (WM_USER + 4)

Private Const TTM_ACTIVATE = WM_USER + 1

Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)

Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)

Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)

Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)

Private Const TTM_SETTITLE = (WM_USER + 32)

Private Const TTS_BALLOON = &H40

Private Const TTS_ALWAYSTIP = &H1

Private Const TTF_SUBCLASS = &H10

Private Const TTF_TRACK = &H20

Private Const TTF_IDISHWND = &H1

Private Const TTM_SETDELAYTIME = (WM_USER + 3)

Private Const TTDT_AUTOPOP = 2

Private Const TTDT_INITIAL = 3

Private Const TTM_TRACKACTIVATE = WM_USER + 17

Private Const TTM_TRACKPOSITION = WM_USER + 18

Private Const WS_POPUP = &H80000000

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hWnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'local variable(s) to hold property value(s)
Private mvarBackColor As Long

Private mvarTitle As String

Private mvarForeColor As Long

Private mvarIcon As ttIconType

Private mvarCentered As Boolean

Private mvarStyle As ttStyleEnum

Private mvarTipText As String

Private mvarVisibleTime As Long

Private mvarDelayTime As Long

Private mvarPopupOnDemand As Boolean

Private m_RightToLeft As Boolean

'private data
Private m_lTTHwnd As Long                                  ' hwnd of the tooltip

Private m_lParentHwnd As Long                              ' hwnd of the window the tooltip attached to

Private ti As TOOLINFO

Public WithEvents m_From As Form
Attribute m_From.VB_VarHelpID = -1

'-------------------------------
Private Const TIME_ONESHOT = 0  'Event occurs once, after uDelay milliseconds.

Private Const TIME_PERIODIC = 1  'Event occurs every uDelay milliseconds.

Private Const TIME_CALLBACK_EVENT_PULSE = &H20  'When the timer expires, Windows calls thePulseEvent function to pulse the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.

Private Const TIME_CALLBACK_EVENT_SET = &H10  'When the timer expires, Windows calls theSetEvent function to set the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.

Private Const TIME_CALLBACK_FUNCTION = &H0   'When the timer expires, Windows calls the function pointed to by the lpTimeProc parameter. This is the default.

Private Declare Function timeKillEvent _
                Lib "winmm.dll" (ByVal uID As Long) As Long

Private Declare Function timeSetEvent _
                Lib "winmm.dll" (ByVal uDelay As Long, _
                                 ByVal uResolution As Long, _
                                 ByVal lpFunction As Long, _
                                 ByVal dwUser As Long, _
                                 ByVal uFlags As Long) As Long

Private VBTimer As Long, MMTimer As Long

Private hMMTimer As Long

Private WithEvents CodeTimer As ImpulseCodeTimer
Attribute CodeTimer.VB_VarHelpID = -1

Private NourTimer As Integer
'-----------------------------------

Private Sub Class_Initialize()
    InitCommonControls
    mvarDelayTime = 500
    mvarVisibleTime = 5000
    mvarPopupOnDemand = False
End Sub

Private Sub Class_Terminate()
    Destroy
    Set CodeTimer = Nothing
End Sub

'//////////////////////////////////////////////////////
Public Property Get VisibleTime() As Long
    VisibleTime = mvarVisibleTime
End Property

Public Property Let VisibleTime(ByVal lData As Long)
    mvarVisibleTime = lData
    Dim LngTemp As Long

    If mvarVisibleTime > 32767 Then mvarVisibleTime = 32767
    If mvarVisibleTime < 0 Then mvarVisibleTime = 0
    
    LngTemp = SendMessageLong(m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime)

End Property

'//////////////////////////////////////////////////////
Public Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property

Public Property Let DelayTime(ByVal lData As Long)
    mvarDelayTime = lData
End Property

'//////////////////////////////////////////////////////
Public Property Let Icon(ByVal vData As ttIconType)
    mvarIcon = vData

    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If

End Property

Public Property Get Icon() As ttIconType
    Icon = mvarIcon
End Property

'//////////////////////////////////////////////////////
Public Property Let ForeColor(ByVal vData As Long)
    mvarForeColor = vData

    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If

End Property

Public Property Get ForeColor() As Long

    ForeColor = mvarForeColor
End Property

'//////////////////////////////////////////////////////
Public Property Let Title(ByVal vData As String)
    mvarTitle = vData

    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If

End Property

Public Property Get Title() As String
    Title = ti.lpStr
End Property

'//////////////////////////////////////////////////////
Public Property Let PopupOnDemand(ByVal vData As Boolean)
    mvarPopupOnDemand = vData
    'If m_lTTHwnd <> 0 Then
    'End If
End Property

Public Property Get PopupOnDemand() As Boolean
    PopupOnDemand = mvarPopupOnDemand
End Property

'//////////////////////////////////////////////////////
Public Property Let BackColor(ByVal vData As Long)
    mvarBackColor = vData

    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If

End Property

Public Property Get BackColor() As Long
    BackColor = mvarBackColor
End Property

'//////////////////////////////////////////////////////
Public Property Let TipText(ByVal vData As String)
    mvarTipText = vData
    ti.lpStr = vData

    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If

End Property

Public Property Get TipText() As String
    TipText = mvarTipText
End Property

'//////////////////////////////////////////////////////
Public Property Let Style(ByVal vData As ttStyleEnum)
    mvarStyle = vData
End Property

Public Property Get Style() As ttStyleEnum
    Style = mvarStyle
End Property

'//////////////////////////////////////////////////////
Public Property Let Centered(ByVal vData As Boolean)
    mvarCentered = vData
End Property

Public Property Get Centered() As Boolean
    Centered = mvarCentered
End Property

'X and Y are in Pixel so dont send vbTwips value
Public Sub Show(Optional x As Long = 0, _
                Optional Y As Long = 0)

    Dim pt As POINTAPI
    Dim ptTip As Long
    Dim ret As Long

    With pt
        .x = x
        .Y = Y
    End With

    ret = ClientToScreen(m_lParentHwnd, pt)

    ptTip = pt.Y * &H10000
    ptTip = ptTip + pt.x

    ' These two messages will set the position of the tooltip:
    ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, ti)
    
    Set CodeTimer = New ImpulseCodeTimer
    CodeTimer.interval = 1000
    CodeTimer.Enabled = True
End Sub

Public Sub TimerProc(ByVal uID As Long, _
                     ByVal uMsg As Long, _
                     ByVal dwUser As Long, _
                     ByVal dw1 As Long, _
                     ByVal dw2 As Long)
    'Increment MMTimer
    MMTimer = MMTimer + 1

    If MMTimer > 5 Then
        Me.Destroy
    End If

End Sub

Public Function CreateToolTip(ByVal ParentHwnd As Long) As Boolean
    Dim lWinStyle As Long
    Dim O As Long
    Dim LngTemp As Long

    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If

    m_lParentHwnd = ParentHwnd

    ''create baloon style if desired
    If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    m_lTTHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, 0&, 0&, 0&, 0&, m_lParentHwnd, 0&, 0&, 0&)

    ''now set our tooltip info structure
    With ti

        '//dont incude TTF_SUBCLASS for on demand
        ''if we want it centered, then set that flag
        If mvarCentered Then
            If mvarPopupOnDemand = False Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
            Else
                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP
            End If

        Else

            If mvarPopupOnDemand = False Then
                .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
            Else
                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_TRANSPARENT
            End If
        End If

        'set the hwnd prop to our parent control's hwnd
        .hWnd = m_lParentHwnd
        .lId = m_lParentHwnd                               '0
        .hInstance = App.hInstance
        '.lpstr = ALREADY SET
        '.lpRect = lpRect
        .lSize = Len(ti)
    End With

    ''add the tooltip structure
    SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti

    ''if we want a title or we want an icon
    If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If

    If mvarForeColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If

    If mvarBackColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If

    LngTemp = SendMessageLong(m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime)
    LngTemp = SendMessageLong(m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime)
    
    If Me.RightToLeft = True Then
        O = GetWindowLong(m_lTTHwnd, gw)
        SetWindowLong m_lTTHwnd, gw, O Or ws
    End If

End Function

Public Sub Destroy()

    If m_lTTHwnd <> 0 Then
        Set CodeTimer = Nothing
        DestroyWindow m_lTTHwnd
    End If

End Sub

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal vNewValue As Boolean)
    m_RightToLeft = vNewValue
End Property

Private Sub CodeTimer_Timer()
    NourTimer = NourTimer + 1

    If NourTimer > 5 Then
        Me.Destroy
    End If

End Sub

Private Sub m_From_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             Y As Single)
    Me.Destroy
End Sub

Private Sub m_From_Paint()
    Me.Destroy
End Sub

Private Sub m_From_Resize()
    Me.Destroy
End Sub

Private Sub m_From_Unload(Cancel As Integer)
    Me.Destroy
End Sub
