VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clstooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
Dim LastNum As Long

Private Declare Function CreateWindowEx _
                Lib "user32" _
                Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal x As Long, _
                                         ByVal Y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private Declare Function GetClientRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long

Private Declare Function DestroyWindow _
                Lib "user32" (ByVal hWnd&) As Long

Private Const WM_USER = &H400

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hWnd        As Long
    uID         As Long
    cRect       As RECT
    hinst       As Long
    lpszText    As String
End Type

Private Const CW_USEDEFAULT = &H80000000

Private m_lngHwnd               As Long

Private m_lngMaxWidth           As Long

Private m_PopupOnDemand As Boolean

Private cRect  As RECT

Private Const TTS_NOPREFIX = &H2

Private Const TTF_TRANSPARENT = &H100

Private Const TTF_CENTERTIP = &H2

Private Const TTF_RIGHTTIP = &H1

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

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private Const WS_POPUP = &H80000000

Private Const TTM_SETDELAYTIME = WM_USER + 3

Private Const TTM_GETDELAYTIME = WM_USER + 21

Private Const TTDT_AUTOMATIC = 0

Private Const TTDT_RESHOW = 1

Private Const TTDT_AUTOPOP = 2

Private Const TTDT_INITIAL = 3

Private Const TTF_IDISHWND = &H1

Private Const TTM_TRACKACTIVATE = WM_USER + 17

Private Const TTM_TRACKPOSITION = WM_USER + 18

Private ti As TOOLINFO

'====================================================================='
' Properties
'====================================================================='

Public Property Get MaxWidth() As Long
    'Set Width = m_lngMaxWidth
End Property

Public Property Let MaxWidth(lngMaxWidth As Long)
    m_lngMaxWidth = lngMaxWidth
    SendMessageLong m_lngHwnd, TTM_SETMAXTIPWIDTH, 0, m_lngMaxWidth
End Property

Public Property Get VisibleTime() As Long
    VisibleTime = SendMessageLong(m_lngHwnd, TTM_GETDELAYTIME, TTDT_AUTOPOP, 0)
End Property

Public Property Let VisibleTime(lngTime As Long)
    Dim LngTemp As Long

    If lngTime > 32767 Then lngTime = 32767
    If lngTime < 0 Then lngTime = 0
    
    LngTemp = SendMessageLong(m_lngHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, lngTime)

End Property

Public Property Get DelayTime() As Long
    DelayTime = SendMessageLong(m_lngHwnd, TTM_GETDELAYTIME, TTDT_INITIAL, 0)
End Property

Public Property Let DelayTime(lngTime As Long)

    If lngTime > 32767 Then lngTime = 32767
    If lngTime < 0 Then lngTime = 0
    SendMessageLong m_lngHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, lngTime
    
End Property

Public Sub Create(lngHwndParent As Long, _
                  sTitle As String, _
                  iIconType As Integer, _
                  lBackColor As Long, _
                  lForeColor As Long, _
                  Optional RTL As Boolean = True)
    '---lngHwndParent   Parent Window
    '---sTitle          Title Of ToolTip Text
    '---iIconType       The Icon Type with Displyed

    Dim O As Long
    Dim lWinStyle As Long

    lWinStyle = lWinStyle Or TTS_BALLOON 'ballon shaped tooltips
    m_lngHwnd = CreateWindowEx(0, "tooltips_class32", 0, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, lngHwndParent, 0, App.hInstance, 0)
    
    '--------------
    AddNewHwnd m_lngHwnd
    '--------------
    
    'setting icon & title
    SendMessage m_lngHwnd, TTM_SETTITLE, iIconType, ByVal sTitle
    
    'setting width
    SendMessageLong m_lngHwnd, TTM_SETMAXTIPWIDTH, 0, m_lngMaxWidth
    
    'setting forecolor
    SendMessage m_lngHwnd, TTM_SETTIPTEXTCOLOR, lForeColor, 0&
    
    'setting backcolor
    SendMessage m_lngHwnd, TTM_SETTIPBKCOLOR, lBackColor, 0&

    'set Right To Left
    If RTL = True Then
        O = GetWindowLong(m_lngHwnd, gw)
        SetWindowLong m_lngHwnd, gw, O Or ws
    End If

End Sub

Public Sub Destroy()
    'Old Code
    'DestroyWindow m_lngHwnd
    '-----------------------
    'New Code in 2/11/2005 By ayman
    Dim i As Integer

    If LastNum > 0 Then
        If UBound(HwndCol) > 0 Then

            For i = 0 To UBound(HwndCol) - 1
                'Destroy or unload all tooltips form the memory
                DestroyWindow HwndCol(i)
            Next i

        End If
    End If

    'Clear the dyanmic array
    Erase HwndCol

End Sub

Public Sub AddControl(ctlTool As Object, _
                      StrCaption As String, _
                      Optional blnCenterTip As Boolean = False)
    'On Error Resume Next
    Dim udtToolInfo As TOOLINFO
    Dim lngRes As Long
    On Error GoTo ErrTrap

    With udtToolInfo
        GetClientRect ctlTool.hWnd, .cRect
        .hWnd = ctlTool.hWnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS

        If blnCenterTip Then
            .uFlags = .uFlags Or TTF_CENTERTIP
            '.uFlags = .uFlags Or TTF_RIGHTTIP Or TTF_TRANSPARENT
            '.uFlags = .uFlags Or ttf_l
        End If

        .uID = ctlTool.hWnd
        .lpszText = StrCaption
        .cbSize = Len(udtToolInfo)
    End With

    lngRes = SendMessage(m_lngHwnd, TTM_ADDTOOLA, 0, udtToolInfo)
    Exit Sub
ErrTrap:
    'Debug.Assert lngRes
End Sub

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

    ret = ClientToScreen(m_lngHwnd, pt)

    ptTip = pt.Y * &H10000
    ptTip = ptTip + pt.x

    ' These two messages will set the position of the tooltip:
    ret = SendMessage(m_lngHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lngHwnd, TTM_TRACKACTIVATE, True, ti)

End Sub

Private Sub Class_Initialize()
    m_lngMaxWidth = 300
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Private Sub AddNewHwnd(NewHwnd As Long)

    If NewHwnd <> 0 Then
        'resize the array to put the new handle
        ReDim Preserve HwndCol(LastNum + 1)
        'Put the New window handle in the array
        HwndCol(UBound(HwndCol) - 1) = NewHwnd
        LastNum = LastNum + 1
    End If

End Sub

Public Property Get PopupOnDemand() As Boolean
    PopupOnDemand = m_PopupOnDemand
End Property

Public Property Let PopupOnDemand(ByVal vNewValue As Boolean)
    m_PopupOnDemand = vNewValue
End Property
