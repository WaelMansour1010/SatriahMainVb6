Attribute VB_Name = "basMsgBoxEx"
' MsgBoxEx.bas
'
' By Herman Liu

Option Explicit

Private Const NV_CLOSEMSGBOX = &H5000&
Private Const NV_MOVEMSGBOX = &H5001&
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, _
    ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private mTitle As String
Private mX As Long
Private mY As Long
Private mPause As Long
Private mHandle As Long


Public Function MsgBoxMove(ByVal hwnd As Long, ByVal inPrompt As String, _
        ByVal inTitle As String, ByVal inButtons As Long, _
               ByVal inX As Long, ByVal inY As Long) As Integer
     mTitle = inTitle: mX = inX:  mY = inY
     SetTimer hwnd, NV_MOVEMSGBOX, 0&, AddressOf NewTimerProc
     MsgBoxMove = MessageBox(hwnd, inPrompt, inTitle, inButtons)
End Function


Public Function MsgBoxPause(ByVal hwnd As Long, ByVal inPrompt As String, _
        ByVal inTitle As String, ByVal inButtons As Long, _
        ByVal inPause As Integer) As Integer
     mTitle = inTitle: mPause = inPause * 1000
     SetTimer hwnd, NV_CLOSEMSGBOX, mPause, AddressOf NewTimerProc
     MsgBoxPause = MessageBox(hwnd, inPrompt, inTitle, inButtons)
End Function


Public Function NewTimerProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wparam As Long, _
        ByVal lparam As Long) As Long
    KillTimer hwnd, wparam
    Select Case wparam
         Case NV_CLOSEMSGBOX
              ' A system class is a window class registered by the system which cannot
              ' be destroyed by a processed, e.g. #32768 (a menu), #32769 (desktop
              ' window), #32770 (dialog box), #32771 (task switch window).
             mHandle = FindWindow("#32770", mTitle)
             If mHandle <> 0 Then
                  SetForegroundWindow mHandle
                  SendKeys "{enter}"
             End If
             
        Case NV_MOVEMSGBOX
             mHandle = FindWindow("#32770", mTitle)
             If mHandle <> 0 Then
                  Dim w As Single, h As Single
                  Dim mBox As RECT
                  w = Screen.Width / Screen.TwipsPerPixelX
                  h = Screen.Height / Screen.TwipsPerPixelY
                  GetWindowRect mHandle, mBox
                  If mX > (w - (mBox.Right - mBox.Left) - 1) Then mX = (w - (mBox.Right - mBox.Left) - 1)
                  If mY > (h - (mBox.Bottom - mBox.Top) - 1) Then mY = (h - (mBox.Bottom - mBox.Top) - 1)
                  If mX < 1 Then mX = 1: If mY < 1 Then mY = 1
                    ' SWP_NOSIZE is to use current size, ignoring 3rd & 4th parameters.
                  SetWindowPos mHandle, HWND_TOPMOST, mX, mY, 0, 0, SWP_NOSIZE
             End If
    End Select
End Function


