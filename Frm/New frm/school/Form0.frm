VERSION 5.00
Begin VB.Form Form0 
   Caption         =   "Form11"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form0.frx":0000
   LinkTopic       =   "Form11"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ScreenWidth As Integer
Dim ScreenHeight As Integer
Dim original_ScreenWidth As Integer
Dim original_ScreenHeight As Integer



Function fix_resolution()
'Code:
'Following Text Boxes is the parameters for the screen resulution
'eg. 800*600
Dim typDevM As typDevMODE
Dim lngResult As Long
Dim intAns    As Integer


' Retrieve info about the current graphics mode
' on the current display device.
lngResult = EnumDisplaySettings(0, 0, typDevM)

' Set the new resolution. Don't change the color
' depth so a restart is not necessary.
With typDevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    .dmPelsWidth = ScreenWidth  'ScreenWidth (640,800,1024, etc)
    .dmPelsHeight = ScreenHeight 'ScreenHeight (480,600,768, etc)
End With

' Change the display settings to the specified graphics mode.
lngResult = ChangeDisplaySettings(typDevM, CDS_TEST)
Select Case lngResult
    Case DISP_CHANGE_RESTART
       ' intAns = MsgBox("You must restart your computer to apply these changes." & _
       '     vbCrLf & vbCrLf & "Do you want to restart now?", _
       '     vbYesNo + vbSystemModal, "Screen Resolution")
       ' If intAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
       ' Call ChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
       ' Message = MsgBox("Screen resolution changed", vbInformation, "Resolution Changed ")
    Case Else
        Message = MsgBox("This Program Cann't Run On Your Computer", vbSystemModal, "Error")
End Select
End Function
Private Sub Form_Unload(Cancel As Integer)
If (original_ScreenWidth <> 1024 Or original_ScreenHeight <> 768) Then
ScreenWidth = original_ScreenWidth
ScreenHeight = original_ScreenHeight
fix_resolution
End If
End Sub


Private Sub Form_Load()
numb = 0
   original_ScreenWidth = Screen.Width \ Screen.TwipsPerPixelX
   original_ScreenHeight = Screen.Height \ Screen.TwipsPerPixelY
   origin_w = Screen.Width \ Screen.TwipsPerPixelX
   origin_w = Screen.Height \ Screen.TwipsPerPixelY
   If (original_ScreenWidth <> 1024 Or original_ScreenHeight <> 768) Then
   ScreenWidth = 1024
    ScreenHeight = 768
   fix_resolution
   End If
Me.Visible = False
INTRO.Show
End Sub


