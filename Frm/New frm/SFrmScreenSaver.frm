VERSION 5.00
Begin VB.Form SFrmScreenSaver 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   ScaleHeight     =   11520
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   1440
   End
   Begin VB.Image Image2 
      Height          =   9045
      Left            =   2400
      Picture         =   "SFrmScreenSaver.frx":0000
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   7245
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   0
      Picture         =   "SFrmScreenSaver.frx":54033
      Top             =   0
      Width           =   3810
   End
End
Attribute VB_Name = "SFrmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Dim rad As Integer
Dim clr As Integer
Dim J As Integer
Dim x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13, x14, x15, y1, y2, y3, y4, y5, y6, y7, y8, y9, y10, y11, y12, y13, y14, y15 As Integer
 
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
I = I + 1
If I > 2 Then
J = J + 1
'frmsalebill1.tmr.Enabled = True
I = 0
J = 0

Unload Me

End If

If KeyAscii = 18 Then Unload Me
End Sub

Private Sub Form_Load()
MakeTopMost (1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
I = I + 1
If I > 2 Then I = 0: J = 0: Unload Me

'frmsalebill1.tmr.Enabled = True
'Unload Me
End Sub

Private Sub Image1_Click()
    With Cmdlg
        '*.jpg,*.jpeg,*.jpe,*.jfif
        .CancelError = False
        .DialogTitle = " ≈Œ Ì«— ’Ê—…"
        'Set The Filter to show pictures only
        .Filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|" & "GIF (*.gif)|*.gif|All Files|*.*" ' choose formats to include
        .ShowOpen
    
        If .FileName <> "" Then
            'Set Me.ImgPic.Picture = LoadPicture(.FileName)
            Me.Picture = LoadPicture(.FileName)
            WebForm.Picture = LoadPicture(.FileName)
            SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", .FileName
        Else

            If Dir(App.path & "\Garphics\wallpaper_Main.jpg") <> "" Then
                Me.Picture = LoadPicture(App.path & "\Garphics\wallpaper_Main.jpg")
                WebForm.Picture = LoadPicture(.FileName)
                SaveSetting StrAppRegPath, "View_Type", "BackGroundImag", App.path & "\Garphics\wallpaper_Main.jpg"
                                
            End If

        End If

    End With

End Sub

 
