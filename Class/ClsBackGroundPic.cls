VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBackGroundPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_Pictrue As IPictureDisp

Public Property Get Picture() As IPictureDisp
    Set Picture = m_Pictrue
End Property

Private Sub Class_Initialize()

    If Dir(App.path & "\Garphics\Grid_WallPaper.jpg") <> "" Then
        Set m_Pictrue = LoadPicture(App.path & "\Garphics\Grid_WallPaper.jpg")
    Else
        Set m_Pictrue = Nothing
    End If

End Sub

Public Property Get SearchWallpaper() As IPictureDisp

    If Dir(App.path & "\Garphics\Search_wallpaper.bmp") <> "" Then
        Set SearchWallpaper = LoadPicture(App.path & "\Garphics\Search_wallpaper.bmp")
    Else
        Set SearchWallpaper = Nothing
    End If

End Property

Public Property Get MoneyWallpaper() As IPictureDisp

    If Dir(App.path & "\Garphics\money.bmp") <> "" Then
        Set MoneyWallpaper = LoadPicture(App.path & "\Garphics\money.bmp")
    Else
        Set MoneyWallpaper = Nothing
    End If

End Property

