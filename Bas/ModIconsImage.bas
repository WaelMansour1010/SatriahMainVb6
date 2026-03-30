Attribute VB_Name = "ModIconsImage"
Option Explicit

Public Declare Function LoadImage _
               Lib "user32" _
               Alias "LoadImageA" (ByVal hinst As Long, _
                                   ByVal lpsz As String, _
                                   ByVal un1 As Long, _
                                   ByVal n1 As Long, _
                                   ByVal n2 As Long, _
                                   ByVal un2 As Long) As Long

Public Const LR_LOADMAP3DCOLORS = &H1000

Public Const LR_CREATEDIBSECTION = &H2000

Public Const LR_LOADFROMFILE = &H10

Public Const LR_LOADTRANSPARENT = &H20

Public Const LR_COPYRETURNORG = &H4

Public Const IMAGE_BITMAP = 0

Public Const IMAGE_ICON = 1

Private Const ILC_COLOR = &H0

Private Const ILC_MASK = &H1

Public Const ILC_COLOR4 = &H4

Public Const ILC_COLOR8 = &H8

Public Const ILC_COLOR16 = &H10

Public Const ILC_COLOR24 = &H18

Public Const ILC_COLOR32 = &H20

Public Const ILD_NORMAL = 0

Public Function LoadIcon(path As String, _
                         cx As Long, _
                         cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
End Function
