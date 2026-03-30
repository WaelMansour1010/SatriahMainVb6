Attribute VB_Name = "BottomWindow"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400

Public OldWindowProc As Long

' Handle messages.
Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const WM_ACTIVATE = &H6
Const WM_PAINT = &HF
Const WM_NCDESTROY = &H82
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const HWND_BOTTOM = 1

Dim lFlags As Long

    If (msg = WM_ACTIVATE) Or (msg = WM_PAINT) Then
        ' We're being activated. Move to bottom.
        lFlags = SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos hwnd, HWND_BOTTOM, _
            0, 0, 0, 0, lFlags
    ElseIf msg = WM_NCDESTROY Then
        ' We're being destroyed.
        ' Restore the original WindowProc.
        SetWindowLong _
            hwnd, GWL_WNDPROC, _
            OldWindowProc
    End If

    NewWindowProc = CallWindowProc( _
        OldWindowProc, hwnd, msg, wParam, _
        lParam)
End Function


